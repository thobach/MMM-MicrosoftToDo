/*
  Node Helper module for MMM-MicrosoftToDo

  Purpose: Microsoft's OAutht 2.0 Token API endpoint does not support CORS,
  therefore we cannot make AJAX calls from the browser without disabling
  webSecurity in Electron.
*/
var NodeHelper = require('node_helper')
const fetch = require('node-fetch')
const Log = require('logger')

module.exports = NodeHelper.create({

  start: function () {
    console.log(this.name + ' helper started ...')
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === 'FETCH_DATA') {
      this.fetchData(payload)
    } else if (notification === 'COMPLETE_TASK') {
      this.completeTask(payload.listId, payload.taskId, payload.config)
    } else {
      console.log(this.name + ' - Did not process event: ' + notification)
    }
  },

  completeTask: function (listId, taskId, config) {
    // copy context to be available inside callbacks
    const self = this

    var patchUrl = `https://graph.microsoft.com/v1.0/me/lists/${listId}/tasks/${taskId}`;

    const updateBody = {
      id: taskId,
      status: 'completed'
    }

    fetch(patchUrl, {
      method: 'PATCH',
      body: JSON.stringify(updateBody),
      headers: {
        'Content-Type': 'application/json',
        Authentication: 'Bearer ' + self.accessToken
      }
    }).then(self.checkFetchStatus)
      .then((response) => {
        self.sendSocketNotification('TASK_COMPLETED_' + config.id)
      })
      .error((error) => self.logError("COMPLETE_TASK_ERROR", error));
  },

  getTodos: function (config) {
    // copy context to be available inside callbacks
    let self = this

    // get access token
    var tokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    var refreshToken = config.oauth2RefreshToken
    const form = new URLSearchParams()
    form.append('client_id', config.oauth2ClientId)
    form.append('scope', 'offline_access user.read ' + (config.completeOnClick ? 'tasks.readwrite' : 'tasks.read'))
    form.append('refresh_token', refreshToken)
    form.append('grant_type', 'refresh_token')
    form.append('client_secret', config.oauth2ClientSecret)

    fetch(tokenUrl, {
      method: 'POST',
      body: form
    }).then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((accessTokenJson) => {
        var accessToken = accessTokenJson.access_token
        self.accessToken = accessToken
        self.fetchList(accessToken, config)
      })
      .catch((error) => { 
        self.logError('FETCH_INFO_ERROR_' + config.id, error)
      });
  },
  fetchList: function (accessToken, config) {
    const self = this
    // get ID of task folder
    var getListUrl = 'https://graph.microsoft.com/v1.0/me/todo/lists/?$top=200'
    fetch(getListUrl, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + accessToken
      }
    }).then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((responseData) => {
        // if list name was provided, retrieve its ID
        if (config.listName !== undefined && config.listName !== '') {
          responseData.value.forEach(element => element.displayName === config.listName ? (config._listId = element.id) : '')
        } else if (config.listId !== undefined && config.listId !== '') {
          // if list ID was provided copy it to internal list ID config and show deprecation warning
          config._listId = config.listId
          console.warn(self.name + ' - Warning, configuration parameter listId is deprecated, please use listName instead, otherwise the module will not work anymore in the future.')
          // TODO: during the next release uncomment the following line to not show the todo list, but the error message instead
          // self.sendSocketNotification('FETCH_INFO_ERROR_' + config.id, {
          // error: 'Config param "listId" is deprecated, use "listName" instead',
          // errorDescription: 'The configuration parameter listId is deprecated, please use listName instead. See https://github.com/thobach/MMM-MicrosoftToDo/blob/master/README.MD#installation' })
        } else {
          // otherwise identify the list ID of the default task list first
          // set listID to default task list "Tasks"
          responseData.value.forEach(element => element.wellknownListName === 'defaultList' ? (config._listId = element.id) : '')
        }

        if (config._listId !== undefined && config._listId !== '') {
          // based on translated configuration data (listName -> listId), get tasks
          self.getTasks(accessToken, config, config._listId)
        } else {
          self.sendSocketNotification('FETCH_INFO_ERROR_' + config.id, { error: '"' + config.listName + '" task folder not found', errorDescription: 'The task folder "' + config.listName + '" could not be found.' })
          console.error(self.name + ' - Error while requesting task folders: Could not find task folder ID for task folder name "' + config.listName + '", or could not find default folder in case no task folder name was provided.')
        }
      }) // function callback for task folders
      .catch(error => self.logError('FETCH_INFO_ERROR_' + config.id, error));
  },
  fetchData: function (config) {
    this.getTodos(config)
  },
  getTasks: function (accessToken, config) {
    const self = this
    var orderBy = (config.orderBy === 'subject' ? '&$orderby=title' : '') + (config.orderBy === 'dueDate' ? '&$orderby=duedatetime/datetime' : '')
    var listUrl = 'https://graph.microsoft.com/v1.0/me/todo/lists/' + config._listId + '/tasks?$top=' + config.itemLimit + '&$filter=status%20ne%20%27completed%27%20and%20duedatetime%2Fdatetime%20gt%20%272021-12-01T00%3A00%3A00%27' + orderBy

    fetch(listUrl, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + accessToken
      }
    }).then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((responseData) => {
        var tasks = responseData.value.map((element) => {
          return {
            id: element.id,
            title: element.title,
            dueDateTime: element.dueDateTime,
            listId: config._listId
          }
        } ) ;

        self.sendSocketNotification('DATA_FETCHED_' + config.id, tasks)
      }) // function callback for task folders
      .catch(self.logError)
  },
  checkFetchStatus: function (response) {
    if (response.ok) {    
      return response
    } else {
      throw Error(response.statusText)
    }
  },
  checkBodyError: function (json) {
    if (json && json.error) {
      throw Error(json.error);
    }
    return json;
  },
  logError: function (notificationName, error) {
    Log.error('[MMM-MicrosoftToDo] - Error fetching access token:' + error);
    this.sendSocketNotification(notificationName, error);
  }
})
