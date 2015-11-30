// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

/// <reference path="../../typings/jquery/jquery.d.ts" />
/// <reference path="../../typings/office-js/office-js.d.ts" />
/// <reference path="home.viewmodel.ts" />
/// <reference path="uri.helper.ts" />


module Controllers {

	export class HomeController {
		vm: ViewModels.HomeViewModel;
		item: any;

		constructor() {
			this.vm = new ViewModels.HomeViewModel();
		}

		initializeHandler(reason) {
			var that = this;

			jQuery(document).ready(function() {
				app.initialize();

				if (jQuery.fn.Dropdown) {
					jQuery('.ms-Dropdown').Dropdown();
				}

				jQuery('#clear').click((evt: JQueryEventObject) => { hc.clearForm(evt) });
				jQuery('#submit').click((evt: JQueryEventObject) => { hc.submitForm(evt) });
			});
		}

		clearForm(evt: JQueryEventObject) {
			this.vm.initialise();
			this.bindModelToView();
		}

		bindModelToView() {
			jQuery('#title').val(this.vm.title);
			jQuery('#name').val(this.vm.name);
			jQuery('#address').val(this.vm.address);
			jQuery('#postcode').val(this.vm.postcode);
			jQuery('#telNumber').val(this.vm.telNumber);
			jQuery('#emailAddress').val(this.vm.emailAddress);
			jQuery('#workSummary').val(this.vm.workSummary);
		}

		bindViewToModel() {
			this.vm.title = parseInt(jQuery('#title').val());
			this.vm.name = jQuery('#name').val();
			this.vm.address = jQuery('#address').val();
			this.vm.postcode = jQuery('#postcode').val();
			this.vm.telNumber = jQuery('#telNumber').val();
			this.vm.emailAddress = jQuery('#emailAddress').val();
			this.vm.workSummary = jQuery('#workSummary').val();
		}


		submitForm(evt: JQueryEventObject) {
			this.bindViewToModel();

			this.item = Office.context.mailbox.item;
			this.setItemBody();
		}

		setItemBody() {
			var that = this;
			var uri = helpers.UriBuilder.buildFromObject("estimates", "1234-1234-1234-1234", this.vm);

			this.item.body.getTypeAsync(
				function(result) {
					if (result.status !== Office.AsyncResultStatus.Failed) {
						if (result.value === Office.MailboxEnums.BodyType.Html) {
							// Body is HTML
							var x = '<a href="' + uri + '">' + uri + '</a>';
							that.item.body.setSelectedDataAsync(
								x,
								{ coercionType: Office.CoercionType.Html },
								function(asyncResult) {
									var x = asyncResult;
								 }
							)
						} else {
							// Body is text
							that.item.body.setSelectedDataAsync(
								uri,
								{ coercionType: Office.CoercionType.Text },
								function(asyncResult) { }
							)
						}
					}
				}
			)
		}
	}

	var hc = new Controllers.HomeController();
	Office.initialize = hc.initializeHandler;

}

