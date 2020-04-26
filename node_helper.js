/*
  Node Helper module for MMM-MicrosoftToDo

  Purpose: Microsoft's OAutht 2.0 Token API endpoint does not support CORS,
  therefore we cannot make AJAX calls from the browser without disabling
  webSecurity in Electron.
*/
var NodeHelper = require('node_helper')
const request = require('request')

module.exports = NodeHelper.create({

  start: function () {
    console.log(this.name + ' helper started ...')
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === 'FETCH_DATA') {
      this.fetchData(payload)
    } else {
      console.log(this.name + ' - Did not process event: ' + notification)
    }
  },

  getTodos: function (config) {
    // copy context to be available inside callbacks
    var self = this

    // get access token
    var tokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token'
    var refreshToken = config.oauth2RefreshToken
    var data = {
      client_id: config.oauth2ClientId,
      scope: 'offline_access user.read tasks.read',
      refresh_token: refreshToken,
      grant_type: 'refresh_token',
      client_secret: config.oauth2ClientSecret
    }
    request.post({
      url: tokenUrl,
      form: data
    },
    function (error, response, body) {
      if (error) {
        console.error(self.name + ' - Error while requesting access token:')
        console.error(error)
      }

      if (body && JSON.parse(body).error) {
        console.error(self.name + ' - Error while requesting access token:')
        console.error(JSON.parse(body))

        self.sendSocketNotification('FETCH_INFO_ERROR_' + config.id, { error: JSON.parse(body).error, errorDescription: JSON.parse(body).error_description })

        return
      }

      const accessTokenJson = JSON.parse(body)
      var accessToken = accessTokenJson.access_token

      // get tasks
      var _getTodos = function () {
        var orderBy = (config.orderBy === 'subject' ? '&$orderby=subject' : '') + (config.orderBy === 'dueDate' ? '&$orderby=duedatetime/datetime' : '')
        var listUrl = 'https://graph.microsoft.com/beta/me/outlook/taskFolders/' + config.listId + '/tasks?$select=subject,status&$top=' + config.itemLimit + '&$filter=status%20ne%20%27completed%27' + orderBy

        request.get({
          url: listUrl,
          headers: {
            Authorization: 'Bearer ' + accessToken
          }
        }, function (error, response, body) {
          if (error) {
            console.error(self.name + ' - Error while requesting access token:')
            console.error(error)
          }

          if (body && JSON.parse(body).error) {
            console.error(self.name + ' - Error while requesting tasks:')
            console.error(JSON.parse(body).error)

            self.sendSocketNotification('FETCH_INFO_ERROR_' + config.id, { error: JSON.parse(body).error.code, errorDescription: JSON.parse(body).error.message })

            return
          }

          // send tasks to front-end
          const tasksJson = JSON.parse(body)
          self.sendSocketNotification('DATA_FETCHED_' + config.id, tasksJson.value)
        })
      }

      // if list ID was provided, retrieve its tasks
      if (config.listId !== undefined && config.listId !== '') {
        _getTodos()
      } else {
        // otherwise identify the list ID of the default task list first
        var taksFoldersUrl = 'https://graph.microsoft.com/beta/me/outlook/taskFolders/?$top=200'

        request.get({
          url: taksFoldersUrl,
          headers: {
            Authorization: 'Bearer ' + accessToken
          }
        }, function (error, response, body) {
          if (error) {
            console.log(self.name + ' - Error while requesting task folders:')
            console.log(error)

            self.sendSocketNotification('FETCH_INFO_ERROR', error)

            return
          }

          // parse response from Microsoft
          var list = JSON.parse(body)

          // set listID to default task list "Tasks"
          list.value.forEach(element => element.isDefaultFolder ? (config.listId = element.id) : '')

          // based on new configuration data (listId), get tasks
          _getTodos()
        } // function callback for task folders
        )
      } // else
    })
  },

  fetchData: function (config) {
    this.getTodos(config)
  }
})
