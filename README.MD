# MMM-MicrosoftToDo 
MMM-MicrosoftToDo is a 3rd party module for [MagicMirror](https://github.com/MichMich/MagicMirror) that shows open tasks from task lists. The tasks are retrieved from Microsoft To Do / Outlook.

The idea of this module is to display a shopping list on your mirror that you can control from your Microsoft To Do app and collaborate on with your family. Originally I used [MMM-Wunderlist](https://github.com/paviro/MMM-Wunderlist), though Wunderlist will be shut down and replaced by Microsoft To Do.

## Functionality
* Display open to do list items from a to do list managed by Microsoft To Do
* Update contents every minute

## Preparation
In order to configure this module, you'll need some configuration parameters from Microsoft, which can be obtained as described in this section. All steps are required as the configuration parameters are needed for the installation below.

### Web App Registration (Client ID)
Register an app at https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps (sign-in first if needed)
1. Click on "+ New registration"
2. Enter as "Name": "Magic Mirror"
3. Select as "Supported account types": "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
4. Select as "Redirect URI": "Web"
5. Set as "Redirect URI": "https://localhost:1234"
6. Click "Register"
7. Go to "Overview" and copy the "Application (client) ID" (Client ID), e.g. "4ef19f40-4892-4905-b999-76041e991f53"


### Client Secret Key Generation (Client Secret)
1. Go to "Certificates & secrets"
2. Click "+ New client secret"
3. Set as "Description": "Magic Mirror"
4. Select as "Expires": "Never"
5. Click "Add"
6. Copy the "Value", e.g. "1Q25ls?TKEDf4RWBKVUbKjnaVu=ytP.[", it will be the client secret you need later

### Authorization Code Generation (Code)
Create authorization code by opening the following URL in your browser (but replace the client_id with your value):
```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=4ef19f40-4892-4905-b999-76041e991f53&scope=offline_access user.read tasks.read&redirect_uri=https://localhost:1234&response_mode=query&state=12345
```
1. Click on "Yes" and wait to be redirected
2. Copy the "code" parameter from the URL in the browser, e.g. "M30cd0dff-af91-d061-8755-ffb3b328aa03" from https://localhost:1234/?code=M30cd0dff-af91-d061-8755-ffb3b328aa03&state=12345

### Refresh Token Generation (Refresh Token)
Generate refresh token to allow the module to retrieve your task list going forward
1. Open Terminal app
2. Run the following command with your "client_id" (Client ID), "code" (Code) and "client_secret" (Client Secret) parameters
```bash
curl -X POST \
  https://login.microsoftonline.com/common/oauth2/v2.0/token \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  -d 'grant_type=authorization_code&client_id=4ef19f40-4892-4905-b999-76041e991f53&scope=offline_access%20user.read%20tasks.read&code=M30cd0dff-af91-d061-8755-ffb3b328aa03&redirect_uri=https%3A%2F%2Flocalhost%3A1234&client_secret=1Q25ls?TKEDf4RWBKVUbKjnaVu=ytP.['
```
3. Copy the value of thte "refresh_token" parameter from the response, e.g. refresh_token: "MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA" from
```javascript
{"token_type":"Bearer","scope":"User.Read Tasks.Read","expires_in":3600,"ext_expires_in":3600,"access_token":"EwBwA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAQMC79sEwx4KyeAHQmWi48AYLyFufrOurKrSQRBVaB1CPFwrn/CniDdPAEQYSzb95gEseUsnH3QAjbfi1s5Ya8oW5tLW2atfMG8ttdCnN/Rl8/ZdzfZcFkFQSgJ3kibm33Ov9NGXkKubHBgk0etwhyBsFR25PcDRZgc9hLIgR3dt9OCTF+Ph5lRQqOX8ABQbYy6zm15BHKSOqbHF82QHntrqUZf9u94J7S0ez0reltsOdrFTK0rWcbTMAWjNQfBLoMvdJnqVgTSF++WvkTQCCjTpu+LDEfA7axWBDfAQIXh+KFh2OfLVhST4NKPQNrRgsyI0Mi3kKak/gC7/are8bxQDZgAACHz3HW3uBQHCQAKbkNX8hKy6+AODSWnD/zP0vkRXsoRkiR3JJH7VxrnB31qo9W3Y9PokxT2SBj/97BPmuiJDJ98j190k+MJ/W1kLsL7PbZN9svwLEk8UpPDF2MDiu7GTdhvS7fPDODRBBrK75fU+n1wynvTas+q3ybFM+d0fiqW0p9eUxb7GsK1A0BwXpARW+oFQ2W1g7lzq/w2ss+NyvbjUpizEWw9hUFgJhi32b3ww/r7f0faZRz9vmsLSahbvywo2HG7gzJRdSFbUnoxXRJdFeQpCfxRIhLbE/fx9rB9e0+ENWRuuWqpuJNFjHW9VcV837KAqEwJE9T0F0bacYIq1mcm/n+7g6HSOs27H9LtEHZec6Z+nPPYLxFiN2uRqKaVC7ZBCf6DLfUtQ9sbaDlzI+UUe+GnCTnc2is/hbjYPtK/leiUvOk3jxYBGuU2Jb1kPxH3ahRFlmBYB0mQtbqgOa9DNi0An9eMy+Y/uaLP7EVweQd2bFcBt7tF1nLMQQNUNBfYUva/b2S4SsTVsbHp2tCRmrufFmLrQtFDaG/zJY+tZmnwy8QjbYch9CwEMWRe8+mrtwWy4JKwK00SxoApoDgsaQK/aRYkR5+kn1N/NSwkByyKK6oMbRlFf9SnRsqdFIlgW7gmeizEPhRGqvwTnltghoQdM/aSqY53UQ2jd6H0vYy1PFBtq2s6EGL0Mf9az7JguOgee/WBrI3bN4hQn/xGnrGpkaghsD/2mtRxOAkoEMdxfYQRO60eknbF2QsyYjoOS6V/9ClhzAg==","refresh_token":"MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA“}
```

## Identification of ID of the task folder (Task List ID)
1. Go to https://to-do.microsoft.com/tasks/?fromOwa=true and sign-in if necessary
2. Select your desired task list
3. Copy the task list ID from the URL, e.g. "AQMkADAwATNiZmYAZC1iODgANS1hNGMyLTAwAi0wMAoALgAAA1Lh_i4gMGJMlZDpJ2d8iWwBAB_qpxIEaN5AgOOJ-T1mRJ8AAzJQqfYAAAA=" from "https://to-do.microsoft.com/tasks/AQMkADAwATNiZmYAZC1iODgANS1hNGMyLTAwAi0wMAoALgAAA1Lh_i4gMGJMlZDpJ2d8iWwBAB_qpxIEaN5AgOOJ-T1mRJ8AAzJQqfYAAAA="


## Installation
1. Navigate into your MagicMirror's `modules` folder and execute `git clone https://github.com/thobach/MMM-MicrosoftToDo.git`.
2. In order to load the MMM-MicrosoftToDo module you will need to add the following configuration to your `config/config.js` file. As `oauth2ClientSecret` you need to provide Client Secret from above, as `oauth2RefreshToken` you set the Refresh Token, as `oauth2ClientId` you set the Client ID and as `listId` you set the Task List ID from above.
```javascript
modules: [
	{
		module: 'MMM-MicrosoftToDo',
		position: 'top_left',	// This can be any of the regions. Best results in left or right regions.
		header: 'Shopping List', // This is optional
		config: {
			oauth2ClientSecret: '1Q25ls%3FTKEDf4RWBKVUbKjnaVu%3DytP.%5B',
			oauth2RefreshToken: 'MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA',
			oauth2ClientId: '4ef19f40-4892-4905-b999-76041e991f53',
			listId: 'AQMkADAwATNiZmYAZC1iODgANS1hNGMyLTAwAi0wMAoALgAAA1Lh_i4gMGJMlZDpJ2d8iWwBAB_qpxIEaN5AgOOJ-T1mRJ8AAzJQqfYAAAA=',
		}
	},
]
```

