/*
  Node Helper module for MMM-MicrosoftToDo

  Purpose: Microsoft's OAutht 2.0 Token API endpoint does not support CORS,
  therefore we cannot make AJAX calls from the browser without disabling
  webSecurity in Electron.
*/
var NodeHelper = require("node_helper");
const fetch = require("node-fetch");
const Log = require("logger");
const { add, formatISO, compareAsc, parseISO } = require("date-fns");
const { RateLimit } = require("async-sema");

module.exports = NodeHelper.create({
  start: function () {
    Log.info(`${this.name} node_helper started ...`);
  },

  socketNotificationReceived: function (notification, payload) {
    if (notification === "FETCH_DATA") {
      this.fetchData(payload);
    } else if (notification === "COMPLETE_TASK") {
      this.completeTask(payload.listId, payload.taskId, payload.config);
    } else {
      Log.warn(`${this.name} - did not process event: ${notification}`);
    }
  },

  completeTask: function (listId, taskId, config) {
    // copy context to be available inside callbacks
    const self = this;

    var patchUrl = `https://graph.microsoft.com/v1.0/me/lists/${listId}/tasks/${taskId}`;

    const updateBody = {
      id: taskId,
      status: "completed"
    };

    fetch(patchUrl, {
      method: "PATCH",
      body: JSON.stringify(updateBody),
      headers: {
        "Content-Type": "application/json",
        Authentication: `Bearer ${self.accessToken}`
      }
    })
      .then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((responseJson) => {
        self.sendSocketNotification(
          `TASK_COMPLETED_${config.id}`,
          responseJson
        );
      })
      .error((error) => self.logError(error));
  },

  getTodos: function (config) {
    // copy context to be available inside callbacks
    const self = this;

    // get access token
    var tokenUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/token";
    var refreshToken = config.oauth2RefreshToken;
    const form = new URLSearchParams();
    form.append("client_id", config.oauth2ClientId);
    form.append(
      "scope",
      "offline_access user.read " +
        (config.completeOnClick ? "tasks.readwrite" : "tasks.read")
    );
    form.append("refresh_token", refreshToken);
    form.append("grant_type", "refresh_token");
    form.append("client_secret", config.oauth2ClientSecret);

    fetch(tokenUrl, {
      method: "POST",
      body: form
    })
      .then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((accessTokenJson) => {
        var accessToken = accessTokenJson.access_token;
        self.accessToken = accessToken;
        self.fetchList(accessToken, config);
      })
      .catch((error) => {
        self.logError(error);
      });
  },
  fetchList: function (accessToken, config) {
    const self = this;

    var filterClause = "";
    const hasListNameInConfig =
      config.listName !== undefined && config.listName !== "";
    // filter by displayName, otherwise, get all the lists
    if (hasListNameInConfig) {
      // Get the list ID based on name
      filterClause = `displayName eq '${config.listName}'`;
    }

    filterClause = encodeURIComponent(filterClause).replaceAll("'", "%27");

    var filter = "";
    if (filterClause !== "") {
      filter = `&$filter=${filterClause}`;
    }

    // get ID of task folder
    var getListUrl = `https://graph.microsoft.com/v1.0/me/todo/lists/?$top=200${filter}`;
    fetch(getListUrl, {
      method: "get",
      headers: {
        Authorization: "Bearer " + accessToken
      }
    })
      .then(self.checkFetchStatus)
      .then((response) => response.json())
      .then(self.checkBodyError)
      .then((responseData) => {
        var listIds = [];
        if (config.plannedTasks.enable) {
          //  Filter out any lists that are in the `includedLists` collection
          listIds = responseData.value
            .filter(
              (list) =>
                config.plannedTasks.includedLists.findIndex(
                  (include) => list.displayName.match(include) !== null
                ) !== -1
            )
            .map((list) => list.id);
        } else if (responseData.value.length > 0) {
          if (!hasListNameInConfig) {
            // If there is no list name in the config and it's not showPlannedTasks, get the default list
            const list = responseData.value.find(
              (element) => element.wellknownListName === "defaultList"
            );
            if (list) {
              listIds.push(list.id);
            }
          } else {
            listIds.push(responseData.value[0].id);
          }
        }

        if (listIds.length > 0) {
          self.getTasks(accessToken, config, listIds);
        } else {
          self.logErrorObject({
            error: `"${config.listName}" task folder not found`,
            errorDescription: `The task folder "${config.listName}" could not be found.`
          });
        }
      }) // function callback for task folders
      .catch((error) => self.logError(error));
  },
  fetchData: function (config) {
    this.getTodos(config);
  },
  getTasks: function (accessToken, config, listIds) {
    const self = this;
    Log.info(
      `[MMM-MicrosoftToDo] - Retrieving Tasks for ${listIds.length} list(s)`
    );

    const limit = RateLimit(2);

    // TODO: Iterate through ALL the lists.  If showplannedtasks, filter out those without
    var promises = listIds.map(async (listId) => {
      const promiseSelf = self;
      var orderBy =
        // sorting by subject is not supported anymore in API v1, hence falling back to created time
        (config.orderBy === "subject" ? "&$orderby=createdDateTime" : "") +
        (config.orderBy === "createdDate" ? "&$orderby=createdDateTime" : "") +
        (config.orderBy === "dueDate" ? "&$orderby=duedatetime/datetime" : "");
      var filterClause = "status ne 'completed'";
      if (config.plannedTasks.enable) {
        var pastDate = formatISO(add(Date.now(), config.plannedTasks.duration));
        filterClause += ` and duedatetime/datetime lt '${pastDate}' and duedatetime/datetime ne null`;
      }

      filterClause = encodeURIComponent(filterClause).replaceAll("'", "%27");
      var listUrl = `https://graph.microsoft.com/v1.0/me/todo/lists/${listId}/tasks?$top=${config.itemLimit}&$filter=${filterClause}${orderBy}`;
      Log.debug(`[MMM-MicrosoftToDo] - Retrieving Tasks ${listUrl}`);
      await limit();
      return fetch(listUrl, {
        method: "get",
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      })
        .then(promiseSelf.checkFetchStatus)
        .then((response) => response.json())
        .then(promiseSelf.checkBodyError)
        .then((responseData) => {
          var tasks = [];
          if (
            responseData.value !== null &&
            responseData.value !== undefined &&
            responseData.value.length > 0
          ) {
            tasks = responseData.value.map((element) => {
              var parsedDate;
              if (element !== undefined && element.dueDateTime !== undefined) {
                parsedDate = parseISO(element.dueDateTime.dateTime);
              }
              return {
                id: element.id,
                title: element.title,
                dueDateTime: element.dueDateTime,
                recurrence: element.recurrence,
                listId: config._listId,
                parsedDate: parsedDate
              };
            });
          }
          return tasks;
        }) // function callback for task folders
        .catch(promiseSelf.logError);
    });

    Log.debug(`[MMM-MicrosoftToDo] - waiting on ${promises.length} promises`);
    Promise.all(promises)
      .then((taskArray) => {
        Log.debug(
          `[MMM-MicrosoftToDo] - processing ${taskArray.length} return values`
        );
        var returnTasks = [];
        taskArray.forEach((element) => {
          if (element !== null && element !== undefined && element.length > 0) {
            element.forEach((task) => returnTasks.push(task));
          }
        });

        returnTasks.sort(self.taskSortCompare);
        if (returnTasks.length > config.itemLimit) {
          returnTasks = returnTasks.slice(0, config.itemLimit - 1);
        }

        Log.debug(
          `[MMM-MicrosoftToDo] - returning ${returnTasks.length} tasks`
        );
        self.sendSocketNotification(`DATA_FETCHED_${config.id}`, returnTasks);
      })
      .catch(self.logError);
  },
  taskSortCompare: function (firstTask, secondTask) {
    if (firstTask.parsedDate === undefined) {
      return 1;
    }
    if (secondTask.parsedDate === undefined) {
      return -1;
    }
    return compareAsc(firstTask.parsedDate, secondTask.parsedDate);
  },

  checkFetchStatus: function (response) {
    if (response.ok) {
      return response;
    } else {
      throw Error(response.statusText);
    }
  },
  checkBodyError: function (json) {
    if (json && json.error) {
      throw Error(json.error);
    }
    return json;
  },
  logError: function (error) {
    Log.error(`[MMM-MicrosoftToDo]: ${error}`);
  },
  logErrorObject: function (errorObject) {
    Log.error(`[MMM-MicrosoftToDo]: ${JSON.stringify(errorObject)}`);
  }
});
