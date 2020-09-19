# Invoke-M365HealthMonitoring
Script for checking M365 health and send degradations to Teams and Telegram.
Use config.json file for configure script credentials, such as:

#Tenant configuration
clientId - ID of AAD registered application
tenantId - ID of your M365 tenant
clientSecret  - Client (application) secret.
graphUrl - default value allready filled

#Time frame configuration (in minutes)
RefreshTime - This value defines of how you schedule a script. If you sheduled it for running every 60 minutes, enter the "60" value.

#Messengers config
teamsWebhookID - Enter here your Teams chat/channel webhook ID. 
telegramBotToken - Required to pre-register telegram bor using @botFather and fill here bot token
telegramChannelId - Enter here chat ID where bot should send notifications. If you want to send notifications to telegram channel, please do not forget to add bot to channel as member with Publisher rights

#Usage example
Invoke-M365HealthMonitoring.ps1 [-Messenger {All | Telegram | Teams}

Work example
![Config_Screen](/images/Readme_image_1.png)  
