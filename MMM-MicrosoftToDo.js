Module.register("MMM-MicrosoftToDo",{

    // Override dom generator.
    getDom: function() {

      // checkbox icon is added based on configuration
      var checkbox = this.config.showCheckbox ? "▢&nbsp; " : "";

      // styled wrapper of the todo list
      var listWrapper = document.createElement("ul");
      listWrapper.style.maxWidth = this.config.maxWidth + 'px';
      listWrapper.style.paddingLeft = '0';
      listWrapper.style.marginTop = '0';
      listWrapper.style.listStyleType = 'none';
      listWrapper.classList.add("small");

      var listItemsText = "";

      // for each entry add styled list items
      if (this.list.length != 0) {
        this.list.forEach(element => listItemsText += "<li style=\"list-style-position:inside; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\">" + checkbox + element.subject + "</li>");
      }
      // otherwise indicate that there are no list entries
      else {
        listItemsText += "<li style=\"list-style-position:inside; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;\">" + this.translate("NO_ENTRIES") + "</li>";
      }

      // add list items to wrapper
      listWrapper.innerHTML = listItemsText;

      return listWrapper;
    },

    getTranslations: function() {
      return {
        en: "translations/en.js",
        de: "translations/de.js"
      }
    },

    socketNotificationReceived: function (notification, payload) {

      if (notification === ("FETCH_INFO_ERROR_" + this.config.id)) {

        console.error('An error occurred while retrieving the todo list from Microsoft To Do. Please check the logs.');
        console.error(payload.error);
        console.error(payload.errorDescription);
        this.list = [ { "subject": "Error occurred: " + payload.error + ". Check logs."} ];

        this.updateDom();

      }

      if (notification === ("DATA_FETCHED_" + this.config.id)) {

        this.list = payload;
        console.log(this.name + ' received list of ' + this.list.length + ' items.');

        // check if module should be hidden according to list size and the module's configuration
        if (this.config.hideIfEmpty) {
          if(this.list.length > 0) {
            if(this.hidden){
              this.show()
            }
          } else {
            if(!this.hidden) {
              console.log(this.name + ' hiding module according to \'hideIfEmpty\' configuration, since there are no tasks present in the list.');
              this.hide()
            }
          }
        }

        this.updateDom();

      }
    },

    start: function() {

      // copy module object to be accessible in callbacks
      var self = this;

      // start with empty list that shows loading indicator
      self.list = [ { subject: this.translate("LOADING_ENTRIES") } ];

      // decide if a module should be shown if todo list is empty
      if(self.config.hideIfEmpty === undefined){
        self.config.hideIfEmpty = false;
      }

      // decide if a checkbox icon should be shown in front of each todo list item
      if(self.config.showCheckbox === undefined){
        self.config.showCheckbox = true;
      }

      // set default max module width
      if(self.config.maxWidth === undefined){
        self.config.maxWidth = '450px';
      }

      // set default limit for number of tasks to be shown
      if(self.config.itemLimit === undefined){
        self.config.itemLimit = '200';
      }

      // in case there are multiple instances of this module, ensure the responses from node_helper are mapped to the correct module
      self.config.id = this.identifier;

      // update tasks every 60s
      var refreshFunction = function(){
        self.sendSocketNotification("FETCH_DATA", self.config);
      }
      refreshFunction();
      setInterval(refreshFunction, 60000);

    },

});
