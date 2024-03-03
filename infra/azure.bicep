@maxLength(50)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

// Azure Bot name
param resourceBaseNameBot string

@description('Required when create Azure Bot service')
param botAadAppClientId string

@secure()
@description('Required by Bot Framework package in your bot project')
param botAadAppClientSecret string

param webAppSKU string

@maxLength(42)
param botDisplayName string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param location string = resourceGroup().location
// SSO
param m365ClientId string
param m365TenantId string
param m365OauthAuthorityHost string
param m365ApplicationIdUri string = 'api://botid-${botAadAppClientId}'
@secure()
param m365ClientSecret string

// ASP ID, whne use it, make sure the ASP provided capabilities which matched web app required. for example "F1" doesn't have alwaysOn=true
param aspId string = '/subscriptions/f93b55c0-59f2-4709-9dc5-33f2bf325782/resourceGroups/WebAppRG/providers/Microsoft.Web/serverfarms/asp-Linux'
// Northwind DB API
param northwindDbApi string
// Debug falg
param debugValue string

// Compute resources for your Web App
// resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
//   kind: 'app'
//   location: location
//   name: serverfarmsName
//   sku: {
//     name: webAppSKU
//   }
// }

// Web App that hosts your bot
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    //serverFarmId: serverfarm.id
    serverFarmId: aspId
    httpsOnly: true
    siteConfig: {
      // alwaysOn: true
      alwaysOn: false
      // specify the Node.js version. without this, the application won't be run
      nodeVersion: '~18'
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure App Service from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~18' // Set NodeJS version to 18.x for your site
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'BOT_ID'
          value: botAadAppClientId
        }
        {
          name: 'BOT_PASSWORD'
          value: botAadAppClientSecret
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseNameBot
    botAadAppClientId: botAadAppClientId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

// SSO
resource webAppSettings 'Microsoft.Web/sites/config@2021-02-01' = {
  name: '${webAppName}/appsettings'
  properties: {
    M365_CLIENT_ID: m365ClientId
    M365_CLIENT_SECRET: m365ClientSecret
    INITIATE_LOGIN_ENDPOINT: uri('https://${webApp.properties.defaultHostName}', 'auth-start.html')
    M365_AUTHORITY_HOST: m365OauthAuthorityHost
    M365_TENANT_ID: m365TenantId
    M365_APPLICATION_ID_URI: m365ApplicationIdUri
    BOT_ID: botAadAppClientId
    BOT_PASSWORD: botAadAppClientSecret
    RUNNING_ON_AZURE: '1'
    NORTHWINDDBAPI_ENDPOINT: northwindDbApi
    DEBUG: debugValue
  }
}

// additional settings
// resource webAppAPIEndpoint 'Microsoft.Web/sites/config@2021-02-01' = {
//   name: '${webAppName}/appsettings'
//   properties: {
//     NORTHWINDDBAPI_ENDPOINT: northwindDbApi
//     DEBUG: debugValue
//   }
// }

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
