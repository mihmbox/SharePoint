//SP.ListOperation.Selection.getSelectedList()



SP.SOD.executeOrDelayUntilScriptLoaded(function () {
	SP.SOD.executeOrDelayUntilScriptLoaded(function () {
		SP.SOD.executeOrDelayUntilScriptLoaded(function () {		
			function getListBaseType() {
				var deferred = jQuery.Deferred();

				var ctx = SP.ClientContext.get_current();
				var web = ctx.get_web();
				var list = web.get_lists().getById(SP.ListOperation.Selection.getSelectedList());
				ctx.load(list);
				ctx.executeQueryAsync(function() {
					deferred.resolve(list.get_baseType());
				}, function (sender, args) {
					deferred.reject(sender, args);
				});

				return deferred.promise();
			}
			
			function getListGuid() {
				var listGuid = '';
				if(SP.ListOperation && SP.ListOperation.Selection) {
					listGuid = SP.ListOperation.Selection.getSelectedList();
				}
				
				return listGuid;
			}
			
			try {
				getListBaseType().then(function(baseType) {
					// Don't do any operations if BaseType != 1 (not Document Library)
					if(baseType != 1) {
						return;
					}
					
					// Check if it's List View
					if(!getListGuid()) {
						return;
					}
					
					// Add Ribbon button
					if (typeof ActiveXObject != 'undefined') {
						// don't add Ribon button if ActiveX is not available (works only in IE)
						SendAttachmentRibbon.init();
					}
					
					// Init User Custom Action
					//SendAttachmentUserCustomAction.deleteUserCustomActionList()
					SendAttachmentUserCustomAction.createCustomAction().done(function(created) {
						if(created) {
							location.reload(true);
						}
					});
				});
			} catch(e) {}
		}, "sp.ribbon.js");
	}, "sp.core.js");
}, "sp.js");

jQuery(function () {
	if (typeof ActiveXObject == 'undefined') {
		// ths class will be used for webkit browsers to hide User CUstom Action because ActiveX is not available
		jQuery('body').addClass('no-activex');
	}
	
	jQuery('head').append('<style>\
		.softcat-sendattachment-img { top: -137px; left: -511px; }\
		.no-activex .ms-core-menu-list [title="Send as attachment"] {display: none !important;}\
	</style>');
});