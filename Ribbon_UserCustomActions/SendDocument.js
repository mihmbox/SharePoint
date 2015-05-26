var SendDocument = (function () {
	// begin Private functions

	var ctx;	
	var fileRefArray;
	var fullUrlForSingleDoc;
	var listItem;
	var notification;
	var selectedFile;
	var selectedListItemCollection;
	var selectedItems;
	var site;
	var web;
	var sdList;

	// Deselects all items in selected list
	function deselectAllItems() {
		var ctxT = GetCurrentCtx();
		var rows = ctxT.clvp.tab.rows;
		DeselectAllItems(ctxT, rows, true);
		UpdateSelectAllCbx(ctxT, false);
	}

	// Gets all selected items
	function onQuerySucceededSelectedList() {
		selectedItems = SP.ListOperation.Selection.getSelectedItems();
		itemsArray = [];
		for (i = 0; i < selectedItems.length; i++) {
			itemsArray.push(sdList.getItemById(selectedItems[i].id));
		}
		
		if(!selectedItems.length) {
			// this happens if action was fired from context menu
			var itemId = parseInt($('.js-callout-mainElement').find('.js-callout-ecbMenu[id]').attr('id'));
			if(itemId) {
				itemsArray.push(sdList.getItemById(itemId));
			}
		}

		for (i = 0; i < itemsArray.length; i++) {
			ctx.load(itemsArray[i], "FileRef", "FileLeafRef", "FSObjType");
		}
		ctx.executeQueryAsync(succeededSelectedItems, onQueryFailed);
	}

	// Builts an array of DocumentsUrls und calls outlook
	function succeededSelectedItems() {

		fileRefArray = new Array(selectedItems.length);
		for (i = 0; i < itemsArray.length; i++) {
			fileRefArray[i] = getRootUrl() + itemsArray[i].get_item('FileRef');
		}
		deselectAllItems();
		getOutlook();
	}

	// Calls outlook and adds the attachments
	function getOutlook() {
		var outlook;
		var otem;
		var modalDialog;
		try {
			//        modalDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose("Send e-mail", "Please, wait...", 300, 400);
			outlook = new ActiveXObject('Outlook.Application');
			otem = outlook.CreateItem(0);
			if (fullUrlForSingleDoc != null) {
				otem.Attachments.Add(fullUrlForSingleDoc);
				if (selectedFile != null) {
					otem.Subject = selectedFile;
				}
				fullUrlForSingleDoc = null;
			} else {
				for (var i = 0; i < fileRefArray.length; i++) {
					otem.Attachments.Add(fileRefArray[i]);
				}
				fileRefArray = null;
				otem.Subject = "Documents from SharePoint document library.";
			}
			otem.Display();
			otem.GetInspector.WindowState = 2;
		} catch (err) {		
			alert("Error: " + err.message + "\n" +
				"The following may have cause this error: \n\n" +
				"1. The Outlook is not installed on the machine.\n" +
				"2. The “Initialize and Script ActiveX controls not marked as safe” option should be selected as “Enable” \n" +
				"Open Internet Explorer, go to Tools, Internet Options.\n" +
				"Click on the security tab and make sure the \"Trusted sites\" zone is selected at the top.\n" +
				"Click on the Sites button.\n" +
				"Add current site to trusted sites and close the window.\n" +
				"Then click on Custom Level at the bottom.\n" +
				"In the following window you will see  settings for “Initialize and Script ActiveX controls not marked as safe”,\n" +
				"select the option “enable” under the same.\n" +
				"3. This function was used not inside of share point domain.\n");
		}
		finally {
			otem = null;
			outlook = null;
			if (notification != null)
				SP.UI.Notify.removeNotification(notification);
			SP.UI.ModalDialog.commonModalDialogClose(1, '0');
		}
	}

	function onQueryFailed(sender, args) {
		alert('Failed ' + args.get_massage());
		if (notification != null) {
			SP.UI.Notify.removeNotification(notification);
		}
	}

	// Prepared variables for obtaining the basic environment(ClientContext, Web, Site)
	function getEnvironmentForSendDocumentsFeature() {
		ctx = SP.ClientContext.get_current();
		web = ctx.get_web();
		site = ctx.get_site();
		ctx.load(site);
		ctx.load(web);
	}

	function getRootUrl() {
		var root;
		if (site.get_serverRelativeUrl() == "/") {
			root = site.get_url();
		} else {
			var index = site.get_url().indexOf(site.get_serverRelativeUrl());
			root = site.get_url().substr(0, index);
		}
		return root;
	}

	// end Private functions
	var SendDocument = function (userEmail) {
		this.userEmail = userEmail;
	}

	/*
	Send documents as attachments
	 */
	SendDocument.prototype.send = function () {
		try {
			var modalDialog = SP.UI.ModalDialog.showWaitScreenWithNoClose("Send e-mail", "Please, wait...", 150, 350);
			setTimeout(function () {
				notification = SP.UI.Notify.addNotification('Outlook loading...', true);
				getEnvironmentForSendDocumentsFeature();
				sdList = web.get_lists().getById(SP.ListOperation.Selection.getSelectedList());
				ctx.load(sdList);
				ctx.executeQueryAsync(onQuerySucceededSelectedList, onQueryFailed);
			}, 0);
		} catch (err) {
			alert("Error: " + err.get_message());
			if (notification != null) {
				SP.UI.Notify.removeNotification(notification);
			}
		}
	}

	return SendDocument;
})();
