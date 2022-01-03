
[![MIT License](https://img.shields.io/apm/l/atomic-design-ui.svg?)](https://github.com/tterb/atomic-design-ui/blob/master/LICENSEs)
[![js-standard-style](https://img.shields.io/badge/code%20style-standard-brightgreen.svg)](http://standardjs.com)
[![Vulnerabilities](https://snyk.io/test/github/thobach/MMM-MicrosoftToDo/badge.svg?)](https://snyk.io/test/github/thobach/MMM-MicrosoftToDo)
[![Build Status](https://travis-ci.com/thobach/MMM-MicrosoftToDo.svg?branch=master)](https://travis-ci.com/thobach/MMM-MicrosoftToDo)
[![Language grade: JavaScript](https://img.shields.io/lgtm/grade/javascript/g/thobach/MMM-MicrosoftToDo.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/thobach/MMM-MicrosoftToDo/context:javascript)
[![Total alerts](https://img.shields.io/lgtm/alerts/g/thobach/MMM-MicrosoftToDo.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/thobach/MMM-MicrosoftToDo/alerts/)

# MMM-MicrosoftToDo
MMM-MicrosoftToDo is a 3rd party module for [MagicMirror](https://github.com/MichMich/MagicMirror) that shows open tasks from task lists and allows to complete them as well. The tasks are retrieved from Microsoft To Do / Outlook.

The idea of this module is to display a shopping list on your mirror that you can control from your Microsoft To Do app and collaborate on with your family. Originally I used [MMM-Wunderlist](https://github.com/paviro/MMM-Wunderlist), though Wunderlist will be shut down and replaced by Microsoft To Do.

View in Microsoft ToDo app:

<img src="/screenshots/microsoft%20todo%20original%20app.png?raw=true" width="200" alt="View in Microsoft ToDo app" />

Screenshot of this MMM-MicrosoftToDo module:

<img src="/screenshots/microsoft%20todo%20magic%20mirror%20module.png?raw=true" width="200" alt="Screenshot of the MMM-MicrosoftToDo module" />


## Functionality
* Display open to do list items from a to do list managed by Microsoft To Do
* Update contents every minute
* Complete task when clicking on it (disabled by default)

## Preparation
In order to configure this module, you'll need some configuration parameters from Microsoft, which can be obtained as described in this section. All steps are required as the configuration parameters are needed for the installation below.

### Web App Registration (Client ID)
Register an app at https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps (sign-in first if needed)
1. Click on "+ New registration"
2. Enter as "Name": "Magic Mirror"
3. Select as "Supported account types": "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
4. Click "Register"
4. On the following screen click on "Add a Redirect URI", then "Add a Platform", select "Web"
5. Set as "Redirect URIs": "https://localhost:1234"
6. Check "Access tokens"
6. Click "Configure"
7. Go to "Overview" and copy the "Application (client) ID" (Client ID), e.g. `4ef19f40-4892-4905-b999-76041e991f53`

Note: This setup does not work under an Azure Germany (with T-Systems as data trustee) account.

### Client Secret Key Generation (Client Secret)
1. Go to "Certificates & secrets"
2. Click "+ New client secret"
3. Set as "Description": "Magic Mirror"
4. Select as "Expires": "Never"
5. Click "Add"
6. Copy the "Value", e.g. `1Q25lsZTKEDf4RWBKVUbKjnaVuXytPaB`, it will be the `client secret` you need later

### Authorization Code Generation (Code)
Create authorization code by opening the following URL in your browser (but replace the client_id with your value):
```
https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=4ef19f40-4892-4905-b999-76041e991f53&scope=offline_access%20user.read%20tasks.readwrite%20tasks.read&redirect_uri=https://localhost:1234&response_mode=query&state=12345
```
1. Login and click on "Yes" and wait to be redirected
2. Copy the `code` parameter from the URL in the browser, e.g. `M30cd0dff-af91-d061-8755-ffb3b328aa03` from
```
https://localhost:1234/?code=M30cd0dff-af91-d061-8755-ffb3b328aa03&state=12345
```

### Refresh Token Generation (Refresh Token)
Generate refresh token to allow the module to retrieve your task list going forward
1. Open Terminal app
2. Run the following command with your `client_id` (Client ID), `code` (Code) and `client_secret` (Client Secret) parameters (don't forget the trailing `'`)
```bash
curl -X POST https://login.microsoftonline.com/common/oauth2/v2.0/token -H "Content-Type: application/x-www-form-urlencoded" -d "grant_type=authorization_code&client_id=4ef19f40-4892-4905-b999-76041e991f53&scope=offline_access%20user.read%20tasks.readwrite%20tasks.read&code=M30cd0dff-af91-d061-8755-ffb3b328aa03&redirect_uri=https%3A%2F%2Flocalhost%3A1234&client_secret=1Q25lsZTKEDf4RWBKVUbKjnaVuXytPaB"
```
3. Copy the value of the `refresh_token` parameter from the response, e.g. `refresh_token`:
```
MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA
```
from the full response
```javascript
{"token_type":"Bearer","scope":"User.Read Tasks.Read","expires_in":3600,"ext_expires_in":3600,"access_token":"EwBwA8l6BAAUO9chh8cJscQLmU+LSWpbnr0vmwwAAQMC79sEwx4KyeAHQmWi48AYLyFufrOurKrSQRBVaB1CPFwrn/CniDdPAEQYSzb95gEseUsnH3QAjbfi1s5Ya8oW5tLW2atfMG8ttdCnN/Rl8/ZdzfZcFkFQSgJ3kibm33Ov9NGXkKubHBgk0etwhyBsFR25PcDRZgc9hLIgR3dt9OCTF+Ph5lRQqOX8ABQbYy6zm15BHKSOqbHF82QHntrqUZf9u94J7S0ez0reltsOdrFTK0rWcbTMAWjNQfBLoMvdJnqVgTSF++WvkTQCCjTpu+LDEfA7axWBDfAQIXh+KFh2OfLVhST4NKPQNrRgsyI0Mi3kKak/gC7/are8bxQDZgAACHz3HW3uBQHCQAKbkNX8hKy6+AODSWnD/zP0vkRXsoRkiR3JJH7VxrnB31qo9W3Y9PokxT2SBj/97BPmuiJDJ98j190k+MJ/W1kLsL7PbZN9svwLEk8UpPDF2MDiu7GTdhvS7fPDODRBBrK75fU+n1wynvTas+q3ybFM+d0fiqW0p9eUxb7GsK1A0BwXpARW+oFQ2W1g7lzq/w2ss+NyvbjUpizEWw9hUFgJhi32b3ww/r7f0faZRz9vmsLSahbvywo2HG7gzJRdSFbUnoxXRJdFeQpCfxRIhLbE/fx9rB9e0+ENWRuuWqpuJNFjHW9VcV837KAqEwJE9T0F0bacYIq1mcm/n+7g6HSOs27H9LtEHZec6Z+nPPYLxFiN2uRqKaVC7ZBCf6DLfUtQ9sbaDlzI+UUe+GnCTnc2is/hbjYPtK/leiUvOk3jxYBGuU2Jb1kPxH3ahRFlmBYB0mQtbqgOa9DNi0An9eMy+Y/uaLP7EVweQd2bFcBt7tF1nLMQQNUNBfYUva/b2S4SsTVsbHp2tCRmrufFmLrQtFDaG/zJY+tZmnwy8QjbYch9CwEMWRe8+mrtwWy4JKwK00SxoApoDgsaQK/aRYkR5+kn1N/NSwkByyKK6oMbRlFf9SnRsqdFIlgW7gmeizEPhRGqvwTnltghoQdM/aSqY53UQ2jd6H0vYy1PFBtq2s6EGL0Mf9az7JguOgee/WBrI3bN4hQn/xGnrGpkaghsD/2mtRxOAkoEMdxfYQRO60eknbF2QsyYjoOS6V/9ClhzAg==","refresh_token":"MCVDmtyumC8ZdlGxrkMwCdur4FbPh7GHjy1gOaOHEA8*c5tTI4oIcJAHU8AflLZMLmcCRWNEa7nMkag2hbipYbZ3QgL9JB1G6Hc*JZY9oA5j46Cq79b6BDx8mivTVen0r*39h2hw9pngKS7MGXjKc9efwayQ7UhsFoCJ2wdhLJzv5gTR2EozgcsKsxHODj!1fV8a1sqVqJ!sU*CBLTMsm0twI2hZqsCkPoSWaMIS5kKFAq7bCsfNaumS11TZYni3XjVwuyHn4DEqfLGjHuSNo1TYRifcG422MdG6bYt7tBxWiw5dlq2iIEMl6O*N2sLWU7W0Ns2*oQbcXIFxRsZ8cHP*6M9dn*9Axg8WbtERj!*jHTfC!Ax6IILXrHDJkdAXgh01PVH!huztVrgd0UC6E4R4zJSOa4ytOKWfQcL9h0NdA“}
```

## Installation
1. Navigate into your MagicMirror's `modules` folder and execute `git clone https://github.com/thobach/MMM-MicrosoftToDo.git`.
2. Install dependencies via: cd MMM-MicrosoftToDo && npm install
3. In order to load the MMM-MicrosoftToDo module you will need to add the following configuration to your `config/config.js` file. As `oauth2ClientSecret` you need to provide Client Secret from above, as `oauth2RefreshToken` you set the Refresh Token, as `oauth2ClientId` you set the Client ID and as `listName` you set the exact name of the task list, or leave it blank to show items from the default "Tasks" list.
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
      listName: 'Shopping List', // optional parameter: if not specified displays tasks from default "Tasks" list, if specified will look for a task list with the specified name (exact spelling)
      // Optional parameter:  see Planned Tasks Configuration
      plannedTasks: {
        enabled: false
      },
      showCheckbox: true, // optional parameter: default value is true and will show a checkbox before each todo list item
      showDueDate: false, // optional parameter: default value is false and will show the todo list items due date if it exists on the todo list item
      dateFormat: 'ddd MMM Do [ - ]', //optional parameter: uses moment date format and the default value is 'ddd MMM Do [ - ]'
      highlightTagColor: '#E3FF30', // optional parameter: highlight tags (#Tags) in the entry text. value can be a HTML color value
      hideIfEmpty: false, // optional parameter: default value is false and will show the module also when the todo list is empty
      maxWidth: 450, // optional parameter: max width in pixel, default value is 450
      itemLimit: 200, // optional parameter: limit on the number of items to show from the list, default value is 200
      orderBy: 'createdDate', // optional parameter: 'createdDate' - order results by creation date, 'dueDate' - order results by due date, default value is unordered, ordering by title is not supported anymore in API version 1
      completeOnClick: true, // optional parameter: default value is false, when set to true complete task when clicking on it
      refreshSeconds: 60, // optional parameter: every how many seconds should the list be updated from the remote service, default value is 60
      fade: true, //optional parameter: default value is false. True will fade the list towards the bottom from the point set in the fadePoint parameter
      fadePoint: 0.5 //optional parameter: decimal value between 0 and 1 sets the point where the fade effect will start,
      colorDueDate: false // optional parameter: default value is false.  True will display colors for overdue (red), upcoming (orange), and future (green) dates
    }
  },
]
```
## Planned Tasks Configuration
Microsoft ToDo offers a Planned smart list which shows tasks with a due date across all lists.  You can use the `plannedTasks` configuration section to mimic this functionality in MMM-MicrosoftToDo.

When enabled, the `listName` parameter is ignored, and tasks are retrieved from all lists which match the following criteria:
* Task Status is not equal to `completed`
* Task DueDate is set
* Task Due date is less than the current date + the configured duration.

In other words, if you just enable the Planned Tasks configuration, the default values for `duration` and `includedLists` will take tasks from all lists which are not completed and have a due date less than 2 weeks from today.  That will include any tasks that are overdue.

### Example
```
{
  module: 'MMM-MicrosoftToDo',
  position: 'bottom_left',	// This can be any of the regions. Best results in left or right regions.
  header: 'Planned Tasks', // This is optional
  config: {
    oauth2ClientSecret: 'yourClientSecret',
    oauth2RefreshToken: 'yourRefreshToken',
    oauth2ClientId: 'yourClientId',
    plannedTasks: {
      enable: true,
      duration: {
        weeks: 2
      },
      includedLists: [ "Legal" ] // string array with regular expression patterns.  If a list's displayName matches one of the patterns, it's included
    }
  }
}
```

### `plannedTasks` options

| Option Name | Type    | Description | Default Value |
| ----------- | ----    | ----------  | ------------- |
| enable      | `bool`  |             | `false`       |
| duration    | `object`| A duration object, as defined in the [date-fns add](https://date-fns.org/v2.27.0/docs/add) function.  | `{ weeks: 2}` |
| includedLists | `string[]`| A string list a collection of RegExp patterns.  If a list's `displayName` matches one of the [RegEx](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/RegExp) patterns, it's included | `[ '.*' ]` |

### Why "Include" instead of "Ignore"?
Microsoft will throttle the number of calls you can make to its APIs, and, the current ToDo API does not allow you to query for tasks across all lists.  Therefore, an API call must be made for each list.  We have noticed throttling during testing with as few as 15 lists, so it is wise to limit the lists you choose to include.  While the default will attempt to use all lists, it's advisable to limit yourself to less than 10 to prevent throttling.
