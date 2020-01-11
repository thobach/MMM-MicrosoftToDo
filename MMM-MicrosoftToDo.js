Module.register("MMM-MicrosoftToDo",{

    // Override dom generator.
    getDom: function() {

      // checkbox icon is added based on configuration
      var checkbox = this.config.showCheckbox ? "▢&nbsp; " : "";

      // styled wrapper of the todo list
      var listWrapper = document.createElement("ul");
      // listWrapper.style.border = 'none';
      // listWrapper.style.width = '100%';
      listWrapper.style.paddingLeft = '0';
      listWrapper.style.marginTop = '0';
      listWrapper.style.listStyleType = 'none';
      //  style=\"list-style-type:none\"
      listWrapper.classList.add("small");

      var listItemsText = "";

      // for each entry add styled list items
      if (this.list.length != 0) {
        this.list.forEach(element => listItemsText += "<li>" + checkbox + element.subject + "</li>");
      }
      // otherwise indicate that there are no list entries
      else {
        listItemsText += "<li style=\"list-style-type:none\">" + this.translate("NO_ENTRIES") + "</li>";
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

      // copy module object to be accessible in callbacks
      var self = this;

      var loadEntriesAndRefresh = function() {

        // Generate access token from refresh token
        var xhttp = new XMLHttpRequest();
        xhttp.open("POST", "https://login.microsoftonline.com/common/oauth2/v2.0/token", true);
        xhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        xhttp.send("grant_type=refresh_token&client_id="+self.config.oauth2ClientId+"&scope=user.read%20tasks.read&refresh_token="+self.config.oauth2RefreshToken+"&client_secret="+self.config.oauth2ClientSecret);
        xhttp.onreadystatechange = function() {

          if (this.readyState == 4 && this.status == 200) {

            accessToken = JSON.parse(this.responseText).access_token;

            // Get task list
            var xhttp = new XMLHttpRequest();
            xhttp.open("GET", "https://graph.microsoft.com/beta/me/outlook/taskFolders/"+self.config.listId+"/tasks?$select=subject,status&$top=200&$filter=status%20ne%20%27completed%27", true);
            xhttp.setRequestHeader("Authorization", "Bearer " + accessToken);
            xhttp.send();
            xhttp.onreadystatechange = function() {

              if (this.readyState == 4 && this.status == 200) {

                // parse response from Microsoft
                var list = JSON.parse(this.responseText);

                // store todo list in module to be used during display (getDom function)
                self.list = list.value;

                self.updateDom();

                if (config.hideIfEmpty) {
                  self.list.value.length > 0 ? self.show() : self.hide();
                }

              } // if readyState

            } // function onreadystatechange

          } // if readyState

        }; // function onreadystatechange

      };

      loadEntriesAndRefresh();

      // refresh the TODO list every 60s
      setInterval(loadEntriesAndRefresh, 10000); //perform every 60 seconds.
    },

});
