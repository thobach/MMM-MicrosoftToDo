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
    Log.info(`${this.name} node_helper started ...`)
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === 'FETCH_DATA') {
      this.fetchData(payload)
    } else if (notification === 'COMPLETE_TASK') {
      this.completeTask(payload.listId, payload.taskId, payload.config)
    } else {
      Log.info(`${this.name} - did not process event: ${notification}`)
    }
  },

  completeTask: function (listId, taskId, config) {
    // copy context to be available inside callbacks
    const self = this

    var patchUrl = `https://graph.microsoft.com/v1.0/me/lists/${listId}/tasks/${taskId}`

    const updateBody = {
      id: taskId,
      status: 'completed'
    }

    fetch(patchUrl, {
      method: 'PATCH',
      body: JSON.stringify(updateBody),
      headers: {
        'Content-Type': 'application/json',
        Authentication: `Bearer ${self.accessToken}`
      }
    }).then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((responseJson) => {
        self.sendSocketNotification(`TASK_COMPLETED_${config.id}`, responseJson)
      })
      .error((error) => self.logError('COMPLETE_TASK_ERROR', error))
  },

  getTodos: function (config) {
    // copy context to be available inside callbacks
    const self = this

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
        self.logError(`FETCH_INFO_ERROR_${config.id}`, error)
      })
  },
  fetchList: function (accessToken, config) {
    const self = this
    
    var filterClause = "wellknownListName eq 'defaultList'";
    if (config.listName !== undefined && config.listName !== '') {
      filterClause = `displayName eq '${config.listName}'`
    }
    
    filterClause = encodeURIComponent(filterClause)

    // get ID of task folder
    var getListUrl = `https://graph.microsoft.com/v1.0/me/todo/lists/?$top=200&$filter=${filterClause}`
    fetch(getListUrl, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + accessToken
      }
    }).then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((responseData) => {

        if (responseData.value.length > 0) {
          config._listId = responseData.value.id
          self.getTasks(accessToken, config, config._listId)
        }
        else {
          self.logError(`FETCH_INFO_ERROR_${config.id}`, { error: `"${config.listName}" task folder not found`, errorDescription: `The task folder "${config.listName}" could not be found.` })
        }
      }) // function callback for task folders
      .catch(error => self.logError(`FETCH_INFO_ERROR_${config.id}`, error))
  },
  fetchData: function (config) {
    this.getTodos(config)
  },
  getTasks: function (accessToken, config) {
    const self = this
    var orderBy = (config.orderBy === 'subject' ? '&$orderby=title' : '') + (config.orderBy === 'dueDate' ? '&$orderby=duedatetime/datetime' : '')
    var filterClause = 'status%20ne%20%27completed%27%20and%20duedatetime%2Fdatetime%20gt%20%272021-12-01T00%3A00%3A00%27'
    var listUrl = `https://graph.microsoft.com/v1.0/me/todo/lists/${config._listId}/tasks?$top=${config.itemLimit}&$filter=${filterClause}${orderBy}`

    fetch(listUrl, {
      method: 'get',
      headers: {
        Authorization: `Bearer ${accessToken}`
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
        })

        self.sendSocketNotification(`DATA_FETCHED_${config.id}`, tasks)
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
      throw Error(json.error)
    }
    return json
  },
  logError: function (notificationName, error) {
    Log.error(`[MMM-MicrosoftToDo] - Error fetching access token: ${error}`)
    this.sendSocketNotification(notificationName, error)
  }
})
