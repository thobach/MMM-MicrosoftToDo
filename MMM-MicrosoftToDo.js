/*
global Module, Log, moment
*/
Module.register('MMM-MicrosoftToDo', {

  // Override dom generator.
  getDom: function () {
    // copy module object to be accessible in callbacks
    var self = this

    // checkbox icon is added based on configuration
    var checkbox = this.config.showCheckbox ? '▢ ' : ''

    // styled wrapper of the todo list
    var listWrapper = document.createElement('ul')
    listWrapper.style.maxWidth = this.config.maxWidth + 'px'
    listWrapper.style.paddingLeft = '0'
    listWrapper.style.marginTop = '0'
    listWrapper.style.listStyleType = 'none'
    listWrapper.classList.add('small')

    // for each entry add styled list items
    if (this.list.length !== 0) {
      // Define variable itemCounter and set to 0
      var itemCounter = 0
      this.list.forEach(function (element) {
        // Get due date array
        var taskDue = ''
        if (self.config.showDueDate === true && element.dueDateTime != null) {
          // timezone is returned as UTC
          taskDue = Object.values(element.dueDateTime)
          // converting time zone to browser provided timezone and formatting time according to configuration
          var taskDueDate = moment.utc(taskDue[0]).tz(Intl.DateTimeFormat().resolvedOptions().timeZone)
          if (self.config.useRelativeDate) {
            taskDue = taskDueDate.fromNow()
          } else {
            taskDue = taskDueDate.format(self.config.dateFormat)
          }
        }
        var listItem = document.createElement('li')
        listItem.style.listStylePosition = 'inside'
        listItem.style.whiteSpace = 'nowrap'
        listItem.style.overflow = 'hidden'
        listItem.style.textOverflow = 'ellipsis'

        // needed for the fade effect
        itemCounter += 1

        // Create fade effect.
        if (self.config.fade && self.config.fadePoint < 1) {
          if (self.config.fadePoint < 0) {
            self.config.fadePoint = 0
          }
          var startingPoint = self.config.itemLimit * self.config.fadePoint
          var steps = self.config.itemLimit - startingPoint
          if (itemCounter >= startingPoint) {
            var currentStep = itemCounter - startingPoint
            listItem.style.opacity = 1 - (1 / steps * currentStep)
          }
        }

        var listItemText = document.createTextNode(checkbox + taskDue + element.title)
        listItem.appendChild(listItemText)
        // complete task when clicked on it
        if (self.config.completeOnClick) {
          listItem.onclick = function () {
            self.sendSocketNotification('COMPLETE_TASK', { module: self.data.identifier, listId: element.listId, taskId: element.id, config: self.config })
          }
        }
        listWrapper.appendChild(listItem)
      })
    } else {
      // otherwise indicate that there are no list entries
      listWrapper.innerHTML += '<li style="list-style-position:inside; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">' + this.translate('NO_ENTRIES') + '</li>'
    }
    return listWrapper
  },

  getTranslations: function () {
    return {
      en: 'translations/en.js',
      de: 'translations/de.js'
    }
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === ('FETCH_INFO_ERROR_' + this.config.id)) {
      Log.error('An error occurred while retrieving the todo list from Microsoft To Do. Please check the logs.')
      Log.error(payload.error)
      Log.error(payload.errorDescription)
      this.list = [{ subject: 'Error occurred: ' + payload.error + '. Check logs.' }]

      this.updateDom()
    }

    if (notification === ('DATA_FETCHED_' + this.config.id)) {
      this.list = payload
      Log.info(this.name + ' received list of ' + this.list.length + ' items.')

      // check if module should be hidden according to list size and the module's configuration
      if (this.config.hideIfEmpty) {
        if (this.list.length > 0) {
          if (this.hidden) {
            this.show()
          }
        } else {
          if (!this.hidden) {
            Log.info(this.name + ' hiding module according to \'hideIfEmpty\' configuration, since there are no tasks present in the list.')
            this.hide()
          }
        }
      }

      this.updateDom()
    }

    if (notification === ('TASK_COMPLETED_' + this.config.id)) {
      this.sendSocketNotification('FETCH_DATA', this.config)
    }
  },

  start: function () {
    // copy module object to be accessible in callbacks
    var self = this

    // start with empty list that shows loading indicator
    self.list = [{ subject: this.translate('LOADING_ENTRIES') }]
    self.validateConfig()

    // update tasks every based on config refresh
    var refreshFunction = function () {
      self.sendSocketNotification('FETCH_DATA', self.config)
    }
    refreshFunction()
    setInterval(refreshFunction, self.config.refreshSeconds * 1000)
  },

  validateConfig: function () {
    var self = this

    // in case there are multiple instances of this module, ensure the responses from node_helper are mapped to the correct module
    self.config.id = this.identifier

    // decide if a module should be shown if todo list is empty
    if (self.config.hideIfEmpty === undefined) {
      self.config.hideIfEmpty = false
    }

    // decide if a checkbox icon should be shown in front of each todo list item
    if (self.config.showCheckbox === undefined) {
      self.config.showCheckbox = true
    }

    // set default max module width
    if (self.config.maxWidth === undefined) {
      self.config.maxWidth = '450px'
    }

    // set default limit for number of tasks to be shown
    if (self.config.itemLimit === undefined) {
      self.config.itemLimit = '200'
    }

    // set default task completion on click to false
    if (self.config.completeOnClick === undefined) {
      self.config.completeOnClick = false
    }

    // decide if the task due date should be shown in front of each todo list item, if it exists
    if (self.config.showDueDate === undefined) {
      self.config.showDueDate = false
    }

    // format to display the due date
    if (self.config.dateFormat === undefined) {
      self.config.dateFormat = 'ddd MMM Do [ - ]'
    }

    // set default refresh interval to 60s
    if (self.config.refreshSeconds === undefined) {
      self.config.refreshSeconds = 60
    }

    // set default useRelativeDate to false
    if (self.config.useRelativeDate === undefined) {
      self.config.useRelativeDate = false
    }

    // set default plannedTasks settings
    if (self.config.plannedTasks === undefined) {
      self.config.plannedTasks = {
        enable: false
      }
    }

    // By default, don't ignore any lists
    if (self.config.plannedTasks.ignoreLists === undefined) {
      self.config.plannedTasks.ignoreLists = []
    }

    // by default, only look at tasks 2 weeks out
    if (self.config.plannedTasks.duration === undefined) {
      self.config.plannedTasks.duration = {
        weeks: 2
      }
    }

    if (self.config.listId !== undefined) {
      Log.error(`${self.name} - configuration parameter listId is invalid, please use listName instead.`)
      return false
    }

    return true
  }

})
