using System;
using System.Linq;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Activation;
using Windows.Media.SpeechRecognition;
using Windows.Storage;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Windows.Security.Credentials;
using System.Threading.Tasks;

namespace CortanaIntegrationWithDynamics365
{
    /// <summary>
    /// Provides application-specific behavior to supplement the default Application class.
    /// </summary>
    sealed partial class App : Application
    {

        // Crm Url
        private const string _clientId = "2b482146-9afb-4273-9aa6-ebb5327da41f";
        //CRM URL
        private const string _resource = "https://cortanadynamics365.crm.dynamics.com";

        private const string RedirectUri = "http://login.net";

        //Azure Directory OAUTH 2.0 AUTHORIZATION ENDPOINT
        private const string _authority = "https://login.windows.net/cortanadynamics365.onmicrosoft.com";

        public static string _accessToken;
        //  public static NavigationService NavigationService { get; private set; }

        private static AuthenticationResult _authResult;

        /// <summary>
        /// Initializes the singleton application object.  This is the first line of authored code
        /// executed, and as such is the logical equivalent of main() or WinMain().
        /// </summary>
        public App()
        {
            this.InitializeComponent();
            this.Suspending += OnSuspending;

        }

        /// <summary>
        /// Invoked when the application is launched normally by the end user.  Other entry points
        /// will be used such as when the application is launched to open a specific file.
        /// </summary>
        /// <param name="e">Details about the launch request and process.</param>
        protected override async void OnLaunched(LaunchActivatedEventArgs e)
        {
#if DEBUG
            if (System.Diagnostics.Debugger.IsAttached)
            {
                this.DebugSettings.EnableFrameRateCounter = true;
            }
#endif
            Frame rootFrame = Window.Current.Content as Frame;

            if (rootFrame == null)
            {
                // Create a Frame to act as the navigation context and navigate to the first page
                rootFrame = new Frame();

                rootFrame.NavigationFailed += OnNavigationFailed;

                if (e.PreviousExecutionState == ApplicationExecutionState.Terminated)
                {
                    //TODO: Load state from previously suspended application
                }

                // Place the frame in the current Window
                Window.Current.Content = rootFrame;
            }

            if (e.PrelaunchActivated == false)
            {
                if (rootFrame.Content == null)
                {
                    rootFrame.Navigate(typeof(MainPage), e.Arguments);
                }
                // Ensure the current window is active
                Window.Current.Activate();

                try
                {
                    StorageFile vcdStorageFile = await Package.Current.InstalledLocation.GetFileAsync(@"Assets\Dynamics365Commands.xml");
                    await Windows.ApplicationModel.VoiceCommands.VoiceCommandDefinitionManager.InstallCommandDefinitionsFromStorageFileAsync(vcdStorageFile);
                    await AuthenticateUser();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Installing Voice Commands Failed: " + ex.ToString());
                }
            }
        }

        private async Task AuthenticateUser()
        {
            ApplicationDataContainer localSettings = ApplicationData.Current.LocalSettings;
            var authContext = new AuthenticationContext(_authority, false);
            try
            {
                _authResult = await authContext.AcquireTokenSilentAsync(_resource, _clientId);
            }
            catch (AggregateException exc)
            {
                // AdalException ex = exc.InnerException as AdalException;
            }
            if (_authResult.Error.Equals("failed_to_acquire_token_silently"))
            {
                _authResult = await authContext.AcquireTokenAsync(_resource, _clientId, new Uri(RedirectUri));//, new PlatformParameters(PromptBehavior.Always, false));
            }

            localSettings.Values["Token"] = _authResult.AccessToken;
            localSettings.Values["Resource"] = _resource;
        }

        /// <summary>
        /// Invoked when Navigation to a certain page fails
        /// </summary>
        /// <param name="sender">The Frame which failed navigation</param>
        /// <param name="e">Details about the navigation failure</param>
        void OnNavigationFailed(object sender, NavigationFailedEventArgs e)
        {
            throw new Exception("Failed to load Page " + e.SourcePageType.FullName);
        }

        /// <summary>
        /// Invoked when application execution is being suspended.  Application state is saved
        /// without knowing whether the application will be terminated or resumed with the contents
        /// of memory still intact.
        /// </summary>
        /// <param name="sender">The source of the suspend request.</param>
        /// <param name="e">Details about the suspend request.</param>
        private void OnSuspending(object sender, SuspendingEventArgs e)
        {
            var deferral = e.SuspendingOperation.GetDeferral();
            //TODO: Save application state and stop any background activity
            deferral.Complete();
        }

        protected async override void OnActivated(IActivatedEventArgs args)
        {

            base.OnActivated(args);

            //Type navigationToPageType;
            //ViewModel.TripVoiceCommand? navigationCommand = null;

            // If the app was launched via a Voice Command, this corresponds to the "show trip to <location>" command. 
            // Protocol activation occurs when a tile is clicked within Cortana (via the background task)
            if (args.Kind == ActivationKind.VoiceCommand || args.Kind == ActivationKind.Search)
            {
                await AuthenticateUser();
                var commandArgs = args as VoiceCommandActivatedEventArgs;

                Windows.Media.SpeechRecognition.SpeechRecognitionResult speechRecognitionResult = commandArgs.Result;

                // Get the name of the voice command and the text spoken. See AdventureWorksCommands.xml for
                // the <Command> tags this can be filled with.
                string voiceCommandName = speechRecognitionResult.RulePath[0];
                string textSpoken = speechRecognitionResult.Text;

                // The commandMode is either "voice" or "text", and it indictes how the voice command
                // was entered by the user.
                // Apps should respect "text" mode by providing feedback in silent form.
                string commandMode = this.SemanticInterpretation("commandMode", speechRecognitionResult);
            }
            // or not the app is already active.
            Frame rootFrame = Window.Current.Content as Frame;

            // Ensure the current window is active
            Window.Current.Activate();
        }

        private string SemanticInterpretation(string interpretationKey, SpeechRecognitionResult speechRecognitionResult)
        {
            return speechRecognitionResult.SemanticInterpretation.Properties[interpretationKey].FirstOrDefault();
        }

    }
}