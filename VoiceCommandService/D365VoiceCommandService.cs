using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Windows.ApplicationModel.AppService;
using Windows.ApplicationModel.Background;
using Windows.ApplicationModel.Resources.Core;
using Windows.ApplicationModel.VoiceCommands;
using Windows.UI.Popups;
using Windows.Media.SpeechRecognition;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace VoiceCommandService
{

    public sealed class D365VoiceCommandService : IBackgroundTask
    {
        /// <summary>
        /// The HTTP client
        /// </summary>
        HttpClient httpClient;

        /// <summary>
        /// the service connection is maintained for the lifetime of a cortana session, once a voice command
        /// has been triggered via Cortana.
        /// </summary>
        VoiceCommandServiceConnection voiceServiceConnection;

        /// <summary>
        /// Lifetime of the background service is controlled via the BackgroundTaskDeferral object, including
        /// registering for cancellation events, signalling end of execution, etc. Cortana may terminate the 
        /// background service task if it loses focus, or the background task takes too long to provide.
        /// 
        /// Background tasks can run for a maximum of 30 seconds.
        /// </summary>
        BackgroundTaskDeferral serviceDeferral;

        /// <summary>
        /// ResourceMap containing localized strings for display in Cortana.
        /// </summary>
        ResourceMap cortanaResourceMap;

        /// <summary>
        /// The context for localized strings.
        /// </summary>
        ResourceContext cortanaContext;

        /// <summary>
        /// Get globalization-aware date formats.
        /// </summary>
        DateTimeFormatInfo dateFormatInfo;
        private const string _authority = "https://login.windows.net/cortanadynamics365.onmicrosoft.com";

        //AuthenticationContext authContext = new AuthenticationContext(_authority, new FileCache());
        public async void Run(IBackgroundTaskInstance taskInstance)
        {
            serviceDeferral = taskInstance.GetDeferral();

            // Register to receive an event if Cortana dismisses the background task. This will
            // occur if the task takes too long to respond, or if Cortana's UI is dismissed.
            // Any pending operations should be cancelled or waited on to clean up where possible.
            taskInstance.Canceled += OnTaskCanceled;

            var triggerDetails = taskInstance.TriggerDetails as AppServiceTriggerDetails;

            // Load localized resources for strings sent to Cortana to be displayed to the user.
            cortanaResourceMap = ResourceManager.Current.MainResourceMap.GetSubtree("Resources");

            // Select the system language, which is what Cortana should be running as.
            cortanaContext = ResourceContext.GetForViewIndependentUse();

            // Get the currently used system date format
            dateFormatInfo = CultureInfo.CurrentCulture.DateTimeFormat;
            // This should match the uap:AppService and VoiceCommandService references from the 
            // package manifest and VCD files, respectively. Make sure we've been launched by
            // a Cortana Voice Command.
            if (triggerDetails != null && triggerDetails.Name == "D365VoiceCommandService")
            {
                try
                {
                    voiceServiceConnection =
                        VoiceCommandServiceConnection.FromAppServiceTriggerDetails(
                            triggerDetails);

                    voiceServiceConnection.VoiceCommandCompleted += OnVoiceCommandCompleted;

                    // GetVoiceCommandAsync establishes initial connection to Cortana, and must be called prior to any 
                    // messages sent to Cortana. Attempting to use ReportSuccessAsync, ReportProgressAsync, etc
                    // prior to calling this will produce undefined behavior.
                    VoiceCommand voiceCommand = await voiceServiceConnection.GetVoiceCommandAsync();

                    string recordName, recordId, entityName, query;
                    Record changeRecord = new Record();

                    // Depending on the operation (defined in Dynamics365Commands.xml)
                    // perform the appropriate command.
                    switch (voiceCommand.CommandName)
                    {
                        case "createRecord":
                            ConnectToCRM();
                            entityName = voiceCommand.Properties["entityname"][0];
                            recordName = this.SemanticInterpretation("recordname", voiceCommand.SpeechRecognitionResult);
                            await CreateRecord(entityName, recordName);
                            break;
                        case "showAllUsers":
                            ConnectToCRM();
                            await GetFullNameSystemUsers();
                            break;
                        case "deleteRecord":
                            ConnectToCRM();
                            entityName = voiceCommand.Properties["entityname"][0];
                            changeRecord.recordName = this.SemanticInterpretation("recordname", voiceCommand.SpeechRecognitionResult);
                            query = GetQuery(entityName, changeRecord.recordName);
                            changeRecord.recordId = await RetrieveMultiple(query, entityName);
                            await Delete(entityName, changeRecord);
                            break;
                        case "updateRecord":
                            ConnectToCRM();
                            recordName = this.SemanticInterpretation("opportunityname", voiceCommand.SpeechRecognitionResult);
                            string fieldDisplayName = voiceCommand.Properties["fieldname"][0];
                            string fieldValue = this.SemanticInterpretation("fieldvalue", voiceCommand.SpeechRecognitionResult);
                            query = "opportunities?$select=name,opportunityid&$filter=name eq '" + recordName + "'";
                            recordId = await RetrieveMultiple(query, "Opportunity");
                            string schemaName = GetSchemaName(fieldDisplayName);
                            await Update(recordId, recordName, schemaName, fieldValue);
                            break;
                        case "showOpportunities":
                            var statusOpportunity = voiceCommand.Properties["status"][0];
                            ConnectToCRM();
                            await GetOpportunities(statusOpportunity);
                            break;
                        case "closeOpportunity":
                            var status = voiceCommand.Properties["status"][0];
                            string opportunityName = this.SemanticInterpretation("opportunityname", voiceCommand.SpeechRecognitionResult);
                            ConnectToCRM();
                            //recordId = await RetrieveMultiple(opportunityName);
                            //if (string.IsNullOrEmpty(recordId))
                            //    ShowException("No record found...");
                            //else
                            await CloseOpportunity(opportunityName, status);
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Error Occurred: " + ex.ToString());
                }
            }
        }

        private string GetQuery(string entityName, string recordName)
        {
            string pluralName = GetPluralName(entityName);
            string query = pluralName + "?$select=";
            switch (pluralName)
            {
                case "opportunities":
                    query += "name,opportunityid&$filter=name";
                    break;
                case "leads":
                    query += "firstname,leadid&$filter=firstname";
                    break;
                case "accounts":
                    query += "name,accountid&$filter=name";
                    break;
                case "contacts":
                    query += "firstname,contactid&$filter=firstname";
                    break;
            }
            query += " eq '" + recordName + "'";
            return query;
        }

        private string GetPluralName(string entityName)
        {
            string pluralName = "";
            switch (entityName)
            {
                case "Opportunity":
                    pluralName = "opportunities";
                    break;
                case "Lead":
                    pluralName = "leads";
                    break;
                case "Account":
                    pluralName = "accounts";
                    break;
                case "Contact":
                    pluralName = "contacts";
                    break;
            }
            return pluralName;
        }

        private async Task CreateRecord(string entityName, string recordName)
        {
            string schemaName = "";
            string pluralName = "";
            switch (entityName)
            {
                case "Opportunity":
                    schemaName = "name";
                    pluralName = "opportunities";
                    break;
                case "Lead":
                    schemaName = "firstname";
                    pluralName = "leads";
                    break;
                case "Account":
                    schemaName = "name";
                    pluralName = "accounts";
                    break;
                case "Contact":
                    schemaName = "firstname";
                    pluralName = "contacts";
                    break;
            }
            JObject record = new JObject();
            record.Add(schemaName, recordName);
            HttpResponseMessage createResponse = await SendAsJsonAsync(httpClient, HttpMethod.Post, pluralName, record); //make generic for other entities
            if (createResponse.IsSuccessStatusCode)
            {
                var userMessage = new VoiceCommandUserMessage();
                userMessage.DisplayMessage = userMessage.SpokenMessage = entityName + " with name " + recordName + " created";
                var response = VoiceCommandResponse.CreateResponse(userMessage);
                await voiceServiceConnection.ReportSuccessAsync(response);
            }
        }

        private string GetSchemaName(string fieldDisplayName)
        {
            string schemaName = "";
            switch (fieldDisplayName)
            {
                case "Description":
                    schemaName = "description";
                    break;
                case "Sales Stage":
                    schemaName = "salesstage";
                    break;
                case "Topic":
                    schemaName = "name";
                    break;
            }
            return schemaName;
        }

        private async Task<string> RetrieveMultiple(string query, string entityName)
        {
            try
            {
                //Retrieve 
                Record selectedRecord = new Record();
                string schemaName = "";
                switch (entityName)
                {
                    case "Opportunity":
                        schemaName = "name";
                        break;
                    case "Lead":
                        schemaName = "firstname";
                        break;
                    case "Account":
                        schemaName = "name";
                        break;
                    case "Contact":
                        schemaName = "firstname";
                        break;
                }
                //"opportunities?$select=name,opportunityid&$filter=name eq '" + name + "'"
                //The URL will change in 2016 to include the API version - api/data/v8.0/systemusers
                HttpResponseMessage retrieveResponse =
                    await httpClient.GetAsync(query);

                if (!retrieveResponse.IsSuccessStatusCode)
                    return null;

                JObject jRetrieveResponse =
                    JObject.Parse(retrieveResponse.Content.ReadAsStringAsync().Result);
                if (jRetrieveResponse["value"].Count() == 0)
                {
                    return string.Empty;
                }
                else
                {
                    var names = new List<Record>();
                    string attr = entityName.ToLower() + "id";
                    //int i = 0;
                    foreach (var data in jRetrieveResponse["value"])
                    {
                        names.Add(new Record { recordId = Convert.ToString(data[attr]), recordName = Convert.ToString(data[schemaName]) });
                    }
                    if (names.Count > 1)
                    {
                        selectedRecord = await DisambiguateRecords(names);
                    }
                    else
                    {
                        selectedRecord = names[0];
                    }
                    return selectedRecord.recordId;
                }
            }
            catch (Exception ex)
            {
                ShowException(ex);
                return null;
            }
        }
        private async Task Delete(string entityName, Record record)
        {
            try
            {
                //Delete
                //The URL will change in 2016 to include the API version - api/data/v8.0/accounts               
                string pluralName = GetPluralName(entityName);
                string responseMessage = "";
                HttpResponseMessage deleteResponse =
                    await httpClient.DeleteAsync(pluralName + "(" + record.recordId + ")");

                if (deleteResponse.IsSuccessStatusCode)
                {
                    responseMessage = entityName + " " + record.recordName + " deleted";
                }
                else
                {
                    responseMessage = entityName + " not found";
                }
                await SendResponse(responseMessage);
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }

        private async Task SendResponse(string responseMessage)
        {
            var userMessage = new VoiceCommandUserMessage();
            userMessage.DisplayMessage = userMessage.SpokenMessage = responseMessage;
            var response = VoiceCommandResponse.CreateResponse(userMessage);
            await voiceServiceConnection.ReportSuccessAsync(response);
        }

        private async Task Update(string opportunityId, string recordName, string schemaName, string fieldValue)
        {
            try
            {
                //Update   
                //opportunity_salesstage            
                if (schemaName.Equals("salesstage"))
                {
                    switch (fieldValue)
                    {
                        case "qualify":
                            fieldValue = "0";
                            break;
                        case "develop":
                            fieldValue = "1";
                            break;
                        case "propose":
                            fieldValue = "2";
                            break;
                        case "close":
                            fieldValue = "3";
                            break;
                    }
                }
                string responseMessage = "";
                JObject opportunity = new JObject
                    {
                        {schemaName, fieldValue}
                    };

                //The URL will change in 2016 to include the API version - api/data/v8.0/accounts
                HttpResponseMessage updateResponse =
                    await SendAsJsonAsync(httpClient, new HttpMethod("PATCH"), "opportunities(" + opportunityId + ")", opportunity);
                if (updateResponse.IsSuccessStatusCode)
                {
                    responseMessage = "Opportunity " + recordName + " updated";
                }
                else
                {
                    responseMessage = "Opportunity not found";
                }
                await SendResponse(responseMessage);
            }
            catch (Exception ex)
            {
                ShowException(ex);
            }
        }
        private async void ShowException(Exception ex)
        {
            var dialog = new MessageDialog(ex.Message);
            await dialog.ShowAsync();

            //Needs to be changed to voice command response.
        }

        private async void ShowException(string message)
        {
            var dialog = new MessageDialog(message);
            await dialog.ShowAsync();

            //Needs to be changed to voice command response.
        }

        /// <summary>
        /// Gets the full name system users.
        /// </summary>
        /// <returns></returns>
        private async Task GetFullNameSystemUsers()
        {
            try
            {

                //Retrieve 
                //The URL will change in 2016 to include the API version - api/data/v8.0/systemusers
                HttpResponseMessage retrieveResponse =
                    await httpClient.GetAsync("systemusers?$select=fullname&$orderby=fullname asc");

                if (retrieveResponse.IsSuccessStatusCode)
                {
                    JObject jRetrieveResponse =
                        JObject.Parse(retrieveResponse.Content.ReadAsStringAsync().Result);
                    dynamic systemUserObject = JsonConvert.DeserializeObject(jRetrieveResponse.ToString());
                    var userMessage = new VoiceCommandUserMessage();
                    var destinationsContentTiles = new List<VoiceCommandContentTile>();
                    string recordsRetrieved = string.Format(
                         cortanaResourceMap.GetValue("UsersRetrieved", cortanaContext).ValueAsString);
                    userMessage.DisplayMessage = userMessage.SpokenMessage = recordsRetrieved;
                    int i = 0;
                    foreach (var data in systemUserObject.value)
                    {
                        if (i >= 10)
                            break;
                        var destinationTile = new VoiceCommandContentTile();
                        destinationTile.ContentTileType = VoiceCommandContentTileType.TitleOnly;
                        destinationTile.Title = data.fullname.Value;
                        destinationsContentTiles.Add(destinationTile);
                        i++;
                    }
                    var response = VoiceCommandResponse.CreateResponse(userMessage, destinationsContentTiles);
                    await voiceServiceConnection.ReportSuccessAsync(response);

                }
            }
            catch (Exception ex)
            {
                //ShowException(ex);
            }
        }

        private async Task GetOpportunities(string status)
        {
            try
            {
                HttpResponseMessage retrieveResponse = null;
                switch (status)
                {
                    case "ACTIVE":
                        retrieveResponse = await httpClient.GetAsync("opportunities?$select=name&$filter=statecode eq 0");
                        break;
                    case "WON":
                        retrieveResponse = await httpClient.GetAsync("opportunities?$select=name&$filter=statuscode eq 3&$orderby=modifiedon desc");
                        break;
                    case "CANCELLED":
                        retrieveResponse = await httpClient.GetAsync("opportunities?$select=name&$filter=statuscode eq 4&$orderby=modifiedon desc");
                        break;
                    case "OUTSOLD":
                        retrieveResponse = await httpClient.GetAsync("opportunities?$select=name&$filter=statuscode eq 5&$orderby=modifiedon desc");
                        break;
                    default:
                        retrieveResponse = await httpClient.GetAsync("opportunities?$select=name&$filter=statecode eq 0");
                        break;
                }

                if (retrieveResponse != null && retrieveResponse.IsSuccessStatusCode)
                {
                    JObject jRetrieveResponse =
                        JObject.Parse(retrieveResponse.Content.ReadAsStringAsync().Result);
                    dynamic systemUserObject = JsonConvert.DeserializeObject(jRetrieveResponse.ToString());
                    var userMessage = new VoiceCommandUserMessage();
                    var destinationsContentTiles = new List<VoiceCommandContentTile>();
                    string recordsRetrieved = string.Format(
                         cortanaResourceMap.GetValue("OpportunitiesRetrieved", cortanaContext).ValueAsString, status);
                    userMessage.DisplayMessage = userMessage.SpokenMessage = recordsRetrieved;
                    int i = 0;
                    foreach (var data in systemUserObject.value)
                    {
                        if (i >= 10)
                            break;
                        var destinationTile = new VoiceCommandContentTile();
                        destinationTile.ContentTileType = VoiceCommandContentTileType.TitleOnly;
                        destinationTile.Title = data.name.Value;
                        destinationsContentTiles.Add(destinationTile);
                        i++;
                    }
                    var response = VoiceCommandResponse.CreateResponse(userMessage, destinationsContentTiles);
                    await voiceServiceConnection.ReportSuccessAsync(response);

                }
            }
            catch (Exception ex)
            {
                //ShowException(ex);
            }
        }

        /// <summary>
        /// Closes the opportunity.
        /// </summary>
        /// <param name="opportunityName">Name of the opportunity.</param>
        /// <param name="status">The status.</param>
        /// <returns></returns>
        private async Task CloseOpportunity(string opportunityName, string status)
        {
            try
            {
                HttpResponseMessage retrieveResponse =
                      await httpClient.GetAsync("opportunities?$select=name,opportunityid&$filter=name eq '" + opportunityName + "' and  statecode eq 0");

                if (retrieveResponse.IsSuccessStatusCode)
                {
                    JObject jRetrieveResponse =
                          JObject.Parse(retrieveResponse.Content.ReadAsStringAsync().Result);

                    dynamic systemUserObject = JsonConvert.DeserializeObject(jRetrieveResponse.ToString());
                    if (jRetrieveResponse["value"].Count() == 0)
                    {
                        var userMessage = new VoiceCommandUserMessage();
                        string recordsRetrieved = string.Format(
                             cortanaResourceMap.GetValue("RecordNotFound", cortanaContext).ValueAsString, opportunityName);
                        userMessage.DisplayMessage = userMessage.SpokenMessage = recordsRetrieved;

                        var response = VoiceCommandResponse.CreateResponse(userMessage);
                        await voiceServiceConnection.ReportSuccessAsync(response);
                        return;
                    }
                    else
                    {
                        Guid id = string.IsNullOrEmpty(Convert.ToString(jRetrieveResponse["value"].First["opportunityid"])) ? Guid.Empty : new Guid(Convert.ToString(jRetrieveResponse["value"].First["opportunityid"]));
                        if (id != Guid.Empty)
                        {
                            if (status == "WON")
                            {
                                JObject opportClose = new JObject();
                                opportClose["subject"] = "Won Opportunity";
                                //The URL will change in 2016 to include the API version - api/data/v8.0/systemusers
                                opportClose["opportunityid@odata.bind"] = httpClient.BaseAddress.ToString() + "opportunities(" + id.ToString() + ")";
                                JObject winOpportParams = new JObject();
                                winOpportParams["Status"] = 3;
                                winOpportParams["OpportunityClose"] = opportClose;
                                retrieveResponse = await SendAsJsonAsync(httpClient, HttpMethod.Post, "WinOpportunity", winOpportParams);
                                if (retrieveResponse.IsSuccessStatusCode)
                                {
                                    var userMessage = new VoiceCommandUserMessage();
                                    string recordsRetrieved = string.Format(
                                         cortanaResourceMap.GetValue("OpportunityClosed", cortanaContext).ValueAsString, opportunityName, status);
                                    userMessage.DisplayMessage = userMessage.SpokenMessage = recordsRetrieved;

                                    var response = VoiceCommandResponse.CreateResponse(userMessage);
                                    await voiceServiceConnection.ReportSuccessAsync(response);
                                }

                            }
                            else if (status == "OUTSOLD")
                            {
                                JObject opportClose = new JObject();
                                opportClose["subject"] = "Lost Opportunity";
                                //The URL will change in 2016 to include the API version - api/data/v8.0/systemusers
                                opportClose["opportunityid@odata.bind"] = httpClient.BaseAddress.ToString() + "opportunities(" + id.ToString() + ")";
                                JObject loseOpportParams = new JObject();
                                loseOpportParams["Status"] = 5;
                                loseOpportParams["OpportunityClose"] = opportClose;
                                retrieveResponse = await SendAsJsonAsync(httpClient, HttpMethod.Post, "LoseOpportunity", loseOpportParams);
                                if (retrieveResponse.IsSuccessStatusCode)
                                {
                                    var userMessage = new VoiceCommandUserMessage();
                                    string recordsRetrieved = string.Format(
                                         cortanaResourceMap.GetValue("OpportunityClosed", cortanaContext).ValueAsString, opportunityName, status);
                                    userMessage.DisplayMessage = userMessage.SpokenMessage = recordsRetrieved;

                                    var response = VoiceCommandResponse.CreateResponse(userMessage);
                                    await voiceServiceConnection.ReportSuccessAsync(response);
                                }
                            }

                            else if (status == "CANCELLED")
                            {
                                JObject opportClose = new JObject();
                                opportClose["subject"] = "Lost Opportunity";
                                //The URL will change in 2016 to include the API version - api/data/v8.0/systemusers
                                opportClose["opportunityid@odata.bind"] = httpClient.BaseAddress.ToString() + "opportunities(" + id.ToString() + ")";
                                JObject loseOpportParams = new JObject();
                                loseOpportParams["Status"] = 4;
                                loseOpportParams["OpportunityClose"] = opportClose;
                                retrieveResponse = await SendAsJsonAsync(httpClient, HttpMethod.Post, "LoseOpportunity", loseOpportParams);
                                if (retrieveResponse.IsSuccessStatusCode)
                                {
                                    var userMessage = new VoiceCommandUserMessage();
                                    string recordsRetrieved = string.Format(
                                         cortanaResourceMap.GetValue("OpportunityClosed", cortanaContext).ValueAsString, opportunityName, status);
                                    userMessage.DisplayMessage = userMessage.SpokenMessage = recordsRetrieved;

                                    var response = VoiceCommandResponse.CreateResponse(userMessage);
                                    await voiceServiceConnection.ReportSuccessAsync(response);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
        }

        /// <summary>
        /// Semantics the interpretation.
        /// </summary>
        /// <param name="interpretationKey">The interpretation key.</param>
        /// <param name="speechRecognitionResult">The speech recognition result.</param>
        /// <returns></returns>
        private string SemanticInterpretation(string interpretationKey, SpeechRecognitionResult speechRecognitionResult)
        {
            return speechRecognitionResult.SemanticInterpretation.Properties[interpretationKey].FirstOrDefault();
        }

        /// <summary>
        /// Sends as json asynchronous.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="client">The client.</param>
        /// <param name="method">The method.</param>
        /// <param name="requestUri">The request URI.</param>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        private async Task<HttpResponseMessage> SendAsJsonAsync<T>(HttpClient client,
          HttpMethod method, string requestUri, T value)
        {
            string content;
            if (value.GetType().Name.Equals("JObject"))
            { content = value.ToString(); }
            else
            {
                content = JsonConvert.SerializeObject(value, new JsonSerializerSettings()
                { DefaultValueHandling = DefaultValueHandling.Ignore });
            }
            HttpRequestMessage request = new HttpRequestMessage(method, requestUri);
            request.Content = new StringContent(content);
            request.Content.Headers.ContentType = MediaTypeHeaderValue.Parse("application/json");
            return await client.SendAsync(request);
        }

        /// <summary>
        /// Show a progress screen. These should be posted at least every 5 seconds for a 
        /// long-running operation, such as accessing network resources over a mobile 
        /// carrier network.
        /// </summary>
        /// <param name="message">The message to display, relating to the task being performed.</param>
        /// <returns></returns>
        private async Task ShowProgressScreen(string message)
        {
            var userProgressMessage = new VoiceCommandUserMessage();
            userProgressMessage.DisplayMessage = userProgressMessage.SpokenMessage = message;

            VoiceCommandResponse response = VoiceCommandResponse.CreateResponse(userProgressMessage);
            await voiceServiceConnection.ReportProgressAsync(response);
        }

        /// <summary>
        /// Provide a simple response that launches the app. Expected to be used in the
        /// case where the voice command could not be recognized (eg, a VCD/code mismatch.)
        /// </summary>
        private async void LaunchAppInForeground()
        {
            var userMessage = new VoiceCommandUserMessage();
            userMessage.SpokenMessage = cortanaResourceMap.GetValue("LaunchingAdventureWorks", cortanaContext).ValueAsString;

            var response = VoiceCommandResponse.CreateResponse(userMessage);

            response.AppLaunchArgument = "";

            await voiceServiceConnection.RequestAppLaunchAsync(response);
        }

        /// <summary>
        /// Handle the completion of the voice command. Your app may be cancelled
        /// for a variety of reasons, such as user cancellation or not providing 
        /// progress to Cortana in a timely fashion. Clean up any pending long-running
        /// operations (eg, network requests).
        /// </summary>
        /// <param name="sender">The voice connection associated with the command.</param>
        /// <param name="args">Contains an Enumeration indicating why the command was terminated.</param>
        private void OnVoiceCommandCompleted(VoiceCommandServiceConnection sender, VoiceCommandCompletedEventArgs args)
        {
            if (this.serviceDeferral != null)
            {
                this.serviceDeferral.Complete();
            }
        }

        private void ConnectToCRM()
        {
            if (httpClient == null)
            {
                var localSettings = Windows.Storage.ApplicationData.Current.LocalSettings;
                httpClient = new HttpClient();
                httpClient.BaseAddress = new Uri(Convert.ToString(localSettings.Values["Resource"]) + "/api/data/v8.1/");
                httpClient.Timeout = new TimeSpan(0, 2, 0);
                httpClient.DefaultRequestHeaders.Add("OData-MaxVersion", "4.0");
                httpClient.DefaultRequestHeaders.Add("OData-Version", "4.0");
                httpClient.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));
                httpClient.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", Convert.ToString(localSettings.Values["Token"]));
            }
        }

        /// <summary>
        /// When the background task is cancelled, clean up/cancel any ongoing long-running operations.
        /// This cancellation notice may not be due to Cortana directly. The voice command connection will
        /// typically already be destroyed by this point and should not be expected to be active.
        /// </summary>
        /// <param name="sender">This background task instance</param>
        /// <param name="reason">Contains an enumeration with the reason for task cancellation</param>
        private void OnTaskCanceled(IBackgroundTaskInstance sender, BackgroundTaskCancellationReason reason)
        {
            System.Diagnostics.Debug.WriteLine("Task cancelled, clean up");
            if (this.serviceDeferral != null)
            {
                //Complete the service deferral
                this.serviceDeferral.Complete();
            }
        }

        /// <summary>
        /// Provide the user with a way to identify which record to select. 
        /// </summary>
        /// <param name="records">The set of records</param>
        private async Task<Record> DisambiguateRecords(IEnumerable<Record> records)
        {
            if (records.Count() > 1)
            {
                // Create the first prompt message.
                var userPrompt = new VoiceCommandUserMessage();
                userPrompt.DisplayMessage =
                    userPrompt.SpokenMessage = "Which record do you want to select?";

                // Create a re-prompt message if the user responds with an out-of-grammar response.
                var userReprompt = new VoiceCommandUserMessage();
                userReprompt.DisplayMessage =
                    userReprompt.SpokenMessage = "Sorry, which one do you want to select?";

                // Create card for each item. 
                var destinationContentTiles = new List<VoiceCommandContentTile>();
                int i = 1;
                foreach (Record record in records)
                {
                    var destinationTile = new VoiceCommandContentTile();

                    destinationTile.ContentTileType = VoiceCommandContentTileType.TitleWith68x68IconAndText;
                    //destinationTile.Image = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///AdventureWorks.VoiceCommands/Images/GreyTile.png"));

                    // The AppContext can be any arbitrary object.
                    destinationTile.AppContext = record;
                    destinationTile.Title = record.recordName;
                    destinationContentTiles.Add(destinationTile);
                    i++;
                }

                // Cortana handles re-prompting if no valid response.
                var response = VoiceCommandResponse.CreateResponseForPrompt(userPrompt, userReprompt, destinationContentTiles);

                // If cortana is dismissed in this operation, null is returned.
                var voiceCommandDisambiguationResult = await
                    voiceServiceConnection.RequestDisambiguationAsync(response);
                if (voiceCommandDisambiguationResult != null)
                {
                    return (Record)voiceCommandDisambiguationResult.SelectedItem.AppContext;
                }
            }
            return null;
        }
    }
}
