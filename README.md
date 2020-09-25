# Invoke-M365HealthMonitoring
Script for checking M365 health and send degradations to Teams and Telegram.  
Pre-requisites:
1. Created AAD "Office 365 Management APIs" Application with app permissions "ServiceHealth.Read"
2. If you want to receive alerts to Teams channel, you need to add Teams Webhook application to the channel firstly  
3. If you want to receive alers to Telegram channel/chat, need to create telegram bot

Use config.json file for configure script credentials, such as:  
![Telegram_Example](/images/Readme_image_2.png)  
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

Messages example (image is clickable)  
[![Telegram_Example](/images/Readme_image_1.png)](https://t.me/M365_Health)
