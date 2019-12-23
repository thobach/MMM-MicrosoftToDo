Module.register("MMM-MicrosoftToDo",{

    // Override dom generator.
    getDom: function() {
    
    	// copy config object to be accessible in callbacks
    	var config = this.config;
        
        // wrapper of the todo list
        var listWrapper = document.createElement("div");
        listWrapper.style.border = 'none';
        listWrapper.style.width = '100%';
        
		// Generate access token from refresh token
		var xhttp = new XMLHttpRequest();
		xhttp.open("POST", "https://login.microsoftonline.com/common/oauth2/v2.0/token", true);
		xhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		xhttp.send("grant_type=refresh_token&client_id="+config.oauth2ClientId+"&scope=user.read%20tasks.read&refresh_token="+config.oauth2RefreshToken+"&client_secret="+config.oauth2ClientSecret);
		xhttp.onreadystatechange = function() {
		
		  if (this.readyState == 4 && this.status == 200) {
		  
			accessToken = JSON.parse(this.responseText).access_token;
			
			// Get task list
			var xhttp = new XMLHttpRequest();
			xhttp.open("GET", "https://graph.microsoft.com/beta/me/outlook/taskFolders/"+config.listId+"/tasks?$select=subject,status&$top=200&$filter=status%20ne%20%27completed%27", true);
			xhttp.setRequestHeader("Authorization", "Bearer " + accessToken);
			xhttp.send();
			xhttp.onreadystatechange = function() {
			
			  if (this.readyState == 4 && this.status == 200) {
			  
				list = JSON.parse(this.responseText);
				listText = "<ul style=\"padding-left:0; margin-top:0;\" class=\"small\">";
				list.value.forEach(element => listText += "<li style=\"list-style-type:none\">▢&nbsp; " + element.subject + "</li>");
				listText += "</ul>";
				listWrapper.innerHTML = listText;
				
			  } // if readyState
			  
			} // function onreadystatechange
			
		  } // if readyState
		  
		}; // function onreadystatechange
        
        return listWrapper;
    },
    
    start: function() {
		var self = this;
		setInterval(function() {
			self.updateDom();
		}, 60000); //perform every 60 seconds.
	},

});
