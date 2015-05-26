// Adds User Custom Action for "Send attachment Action"

var SendAttachmentUserCustomAction = (function () {
	var ACTION_TITLE = 'Send as attachment';

	function getUserCustomActionList() {
		var clientContext = SP.ClientContext.get_current();
		var web  = clientContext.get_web();
		var list = web.get_lists().getById(SP.ListOperation.Selection.getSelectedList());
		collUserCustomAction = list.get_userCustomActions();
		clientContext.load(collUserCustomAction);

		var deferred = jQuery.Deferred();
		clientContext.executeQueryAsync(function () {
			var customActionEnumerator = collUserCustomAction.getEnumerator();
			var foundAction = [];
			while (customActionEnumerator.moveNext()) {
				var oUserCustomAction = customActionEnumerator.get_current();
				if (oUserCustomAction.get_title() == ACTION_TITLE) {
					foundAction.push(oUserCustomAction);
				}
			}
			deferred.resolve(collUserCustomAction, foundAction);
		}, function (sender, args) { //On fail function
			deferred.reject(sender, args);
		});

		return deferred.promise();
	}

	function createUserCustomActionList(actionUrl) {
		var deferred = jQuery.Deferred();
		
		getUserCustomActionList().done(function (allActions, foudActions) {
			if (foudActions.length) {
				deferred.resolve(false);
				return;	
			}
			
			// Create action			
			var oUserCustomAction = collUserCustomAction.add();
			oUserCustomAction.set_location('EditControlBlock');
			oUserCustomAction.set_sequence(0);
			oUserCustomAction.set_title(ACTION_TITLE);
			oUserCustomAction.set_url(actionUrl);
			oUserCustomAction.set_name("SendAttachment.UserAction");

			oUserCustomAction.update();
			
			var clientContext = SP.ClientContext.get_current();
			clientContext.executeQueryAsync(function () {
				deferred.resolve(true);
			}, function (sender, args) { //On fail function
				deferred.reject(sender, args);
			});			
		});
		
		return deferred.promise();
	}

	function deleteUserCustomActionList() {
		var deferred = jQuery.Deferred();		
		
		getUserCustomActionList().done(function (allActions, foudActions) {			
			if (foudActions.length) {	

				var clientContext = SP.ClientContext.get_current();			
				for (var i = 0; i < foudActions.length; i++) {
					foudActions[i].deleteObject();
				}	
				
				clientContext.executeQueryAsync(function () {
					deferred.resolve(true);
				}, function (sender, args) {
					deferred.reject(sender, args);
				});				
			} else {
				deferred.resolve(false);
			}
		});
		
		return deferred.promise();
	}
	return {
		createCustomAction : function() {
			// Create global function 
			var sendDocument = new SendDocument();
			window.SendAttachment_CustomAction = function() {
                if (typeof CalloutManager !== "undefined" && CalloutManager !== null) {
                    CalloutManager.closeAll();
                }            
			
				sendDocument.send();
			}

			return createUserCustomActionList( 'javascript:SendAttachment_CustomAction();');
		},
		deleteUserCustomActionList: deleteUserCustomActionList
	}
})();