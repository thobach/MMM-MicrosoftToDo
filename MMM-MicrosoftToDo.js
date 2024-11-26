/*
global Module, Log, moment
*/
Module.register("MMM-MicrosoftToDo", {
  // Module config defaults.           // Make all changes in your config.js file
  defaults: {
    oauth2ClientSecret: "",
    oauth2RefreshToken: "",
    oauth2ClientId: "",
    orderBy: "dueDate",
    hideIfEmpty: false,
    showCheckbox: true,
    maxWidth: 450,
    itemLimit: 200,
    completeOnClick: false,
    showDueDate: false,
    dateFormat: "ddd MMM Do [ - ]",
    refreshSeconds: 60,
    fade: false,
    fadePoint: 0.5,
    useRelativeDate: false,
    markRecurringItems: true,
    plannedTasks: {
      enable: false,
      includedLists: [".*"], // this is ignored as a default value as the whole
      // 'plannedTasks' object is replaced when property 'enable' is set to true
      duration: {
        weeks: 2 // this is ignored as a default value as the whole
        // 'plannedTasks' object is replaced when property 'enable' is set to
        // true
      }
    },
    colorDueDate: false,
    highlightTagColor: null
  },

  getStyles: function () {
    return ["MMM-MicrosoftToDo.css", "font-awesome.css"];
  },

  // Override dom generator.
  getDom: function () {
    // copy module object to be accessible in callbacks
    var self = this;

    // styled wrapper of the todo list
    var listWrapper = document.createElement("ul");
    listWrapper.style.maxWidth = this.config.maxWidth + "px";
    listWrapper.style.paddingLeft = "0";
    listWrapper.style.marginTop = "0";
    listWrapper.style.listStyleType = "none";
    listWrapper.classList.add("small");

    // for each entry add styled list items
    if (this.list.length !== 0) {
      // Define variable itemCounter and set to 0
      var itemCounter = 0;
      this.list.forEach(function (element) {
        // Get due date array
        var taskDue = "";

        var listSpan = document.createElement("span");
        if (self.config.showCheckbox) {
          listSpan.append(document.createTextNode("▢ "));
        }

        if (self.config.showDueDate === true && element.dueDateTime != null) {
          // timezone is returned as UTC
          taskDue = Object.values(element.dueDateTime);
          // converting time zone to browser provided timezone and formatting time according to configuration
          var taskDueDate = moment
            .utc(taskDue[0])
            .tz(Intl.DateTimeFormat().resolvedOptions().timeZone);

          if (self.config.useRelativeDate) {
            taskDueDate = taskDueDate.add(1, "d"); // Due date in Task defaults to midnight on the day, so add a day to shift due date to midnight the next day
          }

          var classNames = ["mmm-task-due-date"];
          if (self.config.colorDueDate) {
            const now = moment();
            const next24 = moment().add(1, "d");
            // overdue
            if (taskDueDate.isBefore(now)) {
              classNames.push("overdue");
            }

            // due in the next day
            if (taskDueDate.isBetween(now, next24)) {
              classNames.push("soon");
            }

            if (taskDueDate.isAfter(next24)) {
              classNames.push("upcoming");
            }
          }
          if (self.config.useRelativeDate) {
            taskDue = `${taskDueDate.fromNow()} - `;
          } else {
            taskDue = taskDueDate.format(self.config.dateFormat);
          }
          var taskText = document.createElement("i");
          taskText.innerText = taskDue;
          taskText.className = classNames.join(" ");

          listSpan.append(taskText);

          // add icon to recurring items
          if (self.config.markRecurringItems && element.recurrence != null) {
            var recurringIcon = document.createElement("i");
            recurringIcon.className = "fas fa-sync";
            recurringIcon.style = "margin-right:5px;";
            recurringIcon.innerText = " - ";
            listSpan.append(recurringIcon);
          }
        }

        var listItem = document.createElement("li");
        listItem.style.listStylePosition = "inside";
        listItem.style.whiteSpace = "nowrap";
        listItem.style.overflow = "hidden";
        listItem.style.textOverflow = "ellipsis";

        // needed for the fade effect
        itemCounter += 1;

        // Create fade effect.
        if (self.config.fade && self.config.fadePoint < 1) {
          if (self.config.fadePoint < 0) {
            self.config.fadePoint = 0;
          }
          var startingPoint = self.config.itemLimit * self.config.fadePoint;
          var steps = self.config.itemLimit - startingPoint;

          var currentStep = itemCounter - startingPoint;
          if (itemCounter < self.config.itemLimit) {
            listItem.style.opacity = 1 - (1 / steps) * currentStep;
          } else if (itemCounter === self.config.itemLimit) {
            // Set opacity of last item to 90% of the opacity of the second to last item
            listItem.style.opacity =
              0.9 * (1 - (1 / steps) * (currentStep - 1));
          }
        }

        // extract tags (#Tag) from subject an display them differently
        if (element.title) {
          var titleTokens = element.title.match(
            /((#[^\s]+)|(?!\s)[^#]*|\s+)+?/g
          );

          titleTokens.forEach((token) => {
            if (token.startsWith("#")) {
              var tagNode = document.createElement("span");
              tagNode.innerText = token;
              if (self.config.highlightTagColor != null) {
                tagNode.style.color = self.config.highlightTagColor;
              }
              listSpan.append(tagNode);
            } else {
              listSpan.append(document.createTextNode(token));
            }
          });
        }
        listItem.appendChild(listSpan);

        // complete task when clicked on it
        if (self.config.completeOnClick) {
          listItem.onclick = function () {
            self.sendSocketNotification("COMPLETE_TASK", {
              module: self.data.identifier,
              listId: element.listId,
              taskId: element.id,
              config: self.config
            });
          };
        }
        listWrapper.appendChild(listItem);
      });
    } else {
      // otherwise indicate that there are no list entries
      listWrapper.innerHTML +=
        '<li style="list-style-position:inside; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">' +
        this.translate("NO_ENTRIES") +
        "</li>";
    }
    return listWrapper;
  },

  getTranslations: function () {
    return {
      en: "translations/en.js",
      de: "translations/de.js"
    };
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === "FETCH_INFO_ERROR_" + this.config.id) {
      Log.error(
        "An error occurred while retrieving the todo list from Microsoft To Do. Please check the logs."
      );
      Log.error(payload.error);
      Log.error(payload.errorDescription);
      this.list = [
        { subject: "Error occurred: " + payload.error + ". Check logs." }
      ];

      this.updateDom();
    }

    if (notification === "DATA_FETCHED_" + this.config.id) {
      this.list = payload;
      Log.info(this.name + " received list of " + this.list.length + " items.");

      // check if module should be hidden according to list size and the module's configuration
      if (this.config.hideIfEmpty) {
        if (this.list.length > 0) {
          if (this.hidden) {
            this.show();
          }
        } else {
          if (!this.hidden) {
            Log.info(
              this.name +
                " hiding module according to 'hideIfEmpty' configuration, since there are no tasks present in the list."
            );
            this.hide();
          }
        }
      }

      this.updateDom();
    }

    if (notification === "TASK_COMPLETED_" + this.config.id) {
      this.sendSocketNotification("FETCH_DATA", this.config);
    }
  },

  start: function () {
    // copy module object to be accessible in callbacks
    var self = this;

    // start with empty list that shows loading indicator
    self.list = [{ subject: this.translate("LOADING_ENTRIES") }];
    self.validateConfig();

    // update tasks every based on config refresh
    var refreshFunction = function () {
      self.sendSocketNotification("FETCH_DATA", self.config);
    };
    refreshFunction();
    setInterval(refreshFunction, self.config.refreshSeconds * 1000);
  },

  validateConfig: function () {
    var self = this;

    // in case there are multiple instances of this module, ensure the responses from node_helper are mapped to the correct module
    self.config.id = this.identifier;

    if (self.config.listId !== undefined) {
      Log.error(
        `${self.name} - configuration parameter listId is invalid, please use listName instead.`
      );
      return false;
    }

    return true;
  }
});
