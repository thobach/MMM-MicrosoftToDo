Module.register("MMM-MicrosoftToDo",{

    // Override dom generator.
    getDom: function() {

      // checkbox icon is added based on configuration
      var checkbox = self.config.showCheckbox ? "▢&nbsp; " : "";

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

      if (notification === "DATA_FETCHED") {

        this.list = payload;

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

      // start with empty list that shows loading indicator
      this.list = [ { subject: this.translate("LOADING_ENTRIES") } ];

      // decide if a module should be shown if todo list is empty
      if(this.config.hideIfEmpty === undefined){
        this.config.hideIfEmpty = false;
      }

      // decide if a checkbox icon should be shown in front of each todo list item
      if(this.config.showCheckbox === undefined){
        this.config.showCheckbox = true;
      }

      // set default max module width
      if(this.config.maxWidth === undefined){
        this.config.maxWidth = '450px';
      }

      // set default limit for number of tasks to be shown
      if(this.config.itemLimit === undefined){
        this.config.itemLimit = '200';
      }

      // copy module object to be accessible in callbacks
      var self = this;

      // update tasks every 60s
      var refreshFunction = function(){
        self.sendSocketNotification("FETCH_DATA", self.config);
      }
      refreshFunction();
      setInterval(refreshFunction, 5000);

    },

});
