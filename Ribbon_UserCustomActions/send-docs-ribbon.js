// Ribbon button Object for sending documents as attachment
var SendAttachmentRibbon = (function () {
	function getUserEmail() {
		var deferred = jQuery.Deferred();
		var context = SP.ClientContext.get_current();
		var web = context.get_web();
		var currentUser = web.get_currentUser();
		context.load(currentUser);
		context.executeQueryAsync(
			function () {
			var email = currentUser.get_email();
			deferred.resolve(email);
		}, function (sender, args) { //On fail function
			deferred.reject(sender, args);
		});

		return deferred.promise();
	}

	function ribbon_init(pageManager) {
		var ribbon = SP.Ribbon.PageManager.get_instance().get_ribbon();

		createTab(ribbon);
		//    ribbon.setFocusOnRibbon();
		ribbon.pollForStateAndUpdate();
	}

	function createTab(ribbon) {
		// var ribbon = SP.Ribbon.PageManager.get_instance().get_ribbon()
		var files = ribbon.getChild('Ribbon.Document'); //ribbon.getChildByTitle('Files');
		var copies = files.getChild('Ribbon.Documents.Copies');
		if (!copies) {
			return;
		}

		var selectedLayout = copies.get_selectedLayout();

		var section = new CUI.Section(ribbon, 'FDH.SendAttachment.Section', 2, 'Bottom'); //2==OneRow
		selectedLayout.addChildAtIndex(section, 1);

		var controlProperties = new CUI.ControlProperties();
		controlProperties.Command = 'FDH.SendAttachment.Button.Command';
		controlProperties.Id = 'FDH.SendAttachment.ControlProperties';
		controlProperties.TemplateAlias = 'o1';
		controlProperties.ToolTipDescription = 'Send email with attached documents';
		controlProperties.Image32by32 = '/_layouts/15/1033/images/formatmap32x32.png';
		controlProperties.ToolTipTitle = 'Send Documents as attachments';
		controlProperties.LabelText = 'Send as attachments';
		controlProperties.Image32by32Left = -511;
		controlProperties.Image32by32Top = -137;
		//		controlProperties.Image32by32Class = 'FDH-sendattachment-img';

		var button = new CUI.Controls.Button(ribbon, 'FDH.SendAttachment.Button', controlProperties);

		var controlComponent = button.createComponentForDisplayMode('Large');
		var row_1 = section.getRow(1);
		row_1.addChildAtIndex(controlComponent, 0);

		copies.selectLayout(selectedLayout.get_title());
	}

	return {
		init : function () {
			/* Register classes and initialize page component */
			SendAttachmentComponent.registerClass('SendAttachmentComponent', CUI.Page.PageComponent);
			
			// getUserEmail().done(function (userEmail) {
			var pm = SP.Ribbon.PageManager.get_instance();
			pm.add_ribbonInited(ribbon_init);

			var ribbon = null;
			try {
				ribbon = pm.get_ribbon();
			} catch (e) {}

			if (!ribbon) {
				if (typeof(_ribbonStartInit) == "function")
					_ribbonStartInit("Ribbon.Document", false, null);
			} else {
				ribbon_init();
			}

			SendAttachmentComponent.initializePageComponent();
			// });
		}
	}
})();


/* Initialize the page component members */
var SendAttachmentComponent = (function () {
	SendAttachmentComponent = function () {
		SendAttachmentComponent.initializeBase(this);
	}
	SendAttachmentComponent.initializePageComponent = function () {
		var ribbonPageManager = SP.Ribbon.PageManager.get_instance();
		if (null !== ribbonPageManager) {
			var rbnInstance = SendAttachmentComponent.get_instance();
			ribbonPageManager.addPageComponent(rbnInstance);
		}
	}

	SendAttachmentComponent.get_instance = function () {
		if (SendAttachmentComponent.instance == null) {
			SendAttachmentComponent.instance = new SendAttachmentComponent();
		}
		return SendAttachmentComponent.instance;
	}

	SendAttachmentComponent.prototype = {
		/* Create an array of handled commands with handler methods */
		init : function () {
			var sendDocument = new SendDocument();

			this._handledCommands = new Object();
			this._handledCommands['FDH.SendAttachment.Button.Command'] = {
				enable : function () {
					if (typeof ActiveXObject == 'undefined') {
						return false;
					}

					var selectedItems = SP.ListOperation.Selection.getSelectedItems();
					for (var i = 0; i < selectedItems.length; i++) {
						if (selectedItems[i].fsObjType == 1) {
							return false;
						}
					}
					return selectedItems.length > 0;
				},
				handle : function (commandId, props, seq) {
					// var listItemId = parseInt(SP.ListOperation.Selection.getSelectedItems()[0].id);
					// var listId = _spPageContextInfo.pageListId;
					sendDocument.send();
				}
			};
			this._commands = ['FDH.SendAttachment.Button.Command'];
		},

		getFocusedCommands : function () {
			return [];
		},
		getGlobalCommands : function () {
			return this._commands;
		},
		canHandleCommand : function (commandId) {
			var handlerFn = this._handledCommands[commandId].enable;
			if (typeof(handlerFn) == 'function') {
				return handlerFn();
			}
			return true;
		},
		handleCommand : function (commandId, properties, sequence) {
			return this._handledCommands[commandId].handle(commandId, properties, sequence);
		},
		isFocusable : function () {
			return false;
		},
		yieldFocus : function () {
			return false;
		},
		receiveFocus : function () {
			return true;
		},
		handleGroup : function () {}
	}

	return SendAttachmentComponent;
})();