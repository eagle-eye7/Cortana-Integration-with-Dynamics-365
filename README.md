# Cortana-Integration-with-Dynamics-365
Initial Commit

Deployment Guide for Cortana Integration with Dynamics 365  
 




Prepared for
18-Apr-17
Version 1.0 Draft

Prepared by
•	Mohammed Farhan Ali 





Revision and Signoff Sheet
Change Record
Date	Author	Version	Change Reference
			
Reviewers
Name	Version Approved	Position	Date
			
Approvers
Name	Version Approved	Position	Date
			

 
Table of Contents
Document Purpose	4
1	Introduction	5
2	Project Objectives	6
3	Deployment Steps	7
 

 

Document Purpose 
This document provides a step by step guide for deployment of Cortana Integration with Dynamics 365.


 
1	Introduction 
•	Cortana integration is available as a Preview feature to organizations that use CRM Online 2016 Update. 
•	It helps users keep track of important Dynamics 365 activities and records such sales activities, accounts, and opportunities which are relevant to salespeople.
•	This project extends Cortana with features and functionality from windows app (as a background task) using voice commands that specify an action or command to execute in CRM.
•	When a voice command is provided, app handles voice command in the background and connects to CRM. After the operation in CRM, it returns all feedback and results through the Cortana canvas and the Cortana voice.

2	Project Objectives
•	Create an app that Cortana invokes in the background.
•	Create an XML document that defines all the spoken commands that the user can say to initiate actions or invoke commands when activating the app.
•	Register the command sets in the XML file when the app is launched.
•	Handle the background activation of the app service as well as the execution of the voice command.
•	Display and speak the appropriate feedback to the voice command within Cortana.





3	Deployment Steps 
Pre-Requisites:
•	Ensure Cortana is signed in with an MSA account. This can be achieved by opening Cortana once and following the sign-in process.
Steps:

1.	Download the solution from the link https://github.com/eagle-eye7/Cortana-Integration-with-Dynamics-365.git 
2.	Create an account in Azure Active Directory and copy the authority and client id, to be placed in app.xaml.cs
3.	Build the solution.
4.	 Run the App at least once to start the services in the background.
5.	Open Cortana and speak the command, such as, “Dynamics Create Account named Test” to create a sample account in CRM. 
6.	Similarly try out other commands like “Show all users” or “Update record”. These are defined the VCD file present in the solution (Dynamics365Commands.xml).
