// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

/// <reference path="../../typings/jquery/jquery.d.ts" />
/// <reference path="../../typings/office-js/office-js.d.ts" />
/// <reference path="home.viewmodel.ts" />
/// <reference path="uri.helper.ts" />


module Controllers {

	export class HomeController {
		vm: ViewModels.HomeViewModel;
		item: Office.Types.MessageCompose;

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

				hc.clearForm(null);
				hc.item = Office.cast.item.toMessageCompose(Office.context.mailbox.item);
			});
		}

		clearForm(evt: JQueryEventObject) {
			this.vm.initialise();
			this.vm.fromEmailAddress = Office.context.mailbox.userProfile.emailAddress;
			this.bindModelToView();
		}

		bindModelToView() {
			jQuery('#title').val(this.vm.title);
			jQuery('#name').val(this.vm.name);
			jQuery('#address').val(this.vm.address);
			jQuery('#postcode').val(this.vm.postcode);
			jQuery('#telNumber').val(this.vm.telNumber);
			jQuery('#emailAddress').val(this.vm.fromEmailAddress);
			jQuery('#workSummary').val(this.vm.workSummary);
		}

		bindViewToModel() {
			var that = this;

			this.vm.title = parseInt(jQuery('#title').val());
			this.vm.name = jQuery('#name').val();
			this.vm.address = jQuery('#address').val();
			this.vm.postcode = jQuery('#postcode').val();
			this.vm.telNumber = jQuery('#telNumber').val();
			this.vm.fromEmailAddress = jQuery('#emailAddress').val();
			this.vm.workSummary = jQuery('#workSummary').val();

			this.item.to.getAsync(
				function(asyncResult) {
					that.vm.toEmailAddresses = asyncResult.value
						.map(function(x) {
							return x.emailAddress;
						});
				});
		}


		submitForm(evt: JQueryEventObject) {
			this.bindViewToModel();
			this.setItemBody();
		}

		setItemBody() {
			var that = this;
			var uri = helpers.UriBuilder.buildFromObject("est", "1234-1234-1234-1234", this.vm);
			var embedString = "";

			this.item.body.getTypeAsync(
				function(result) {
					if (result.status !== Office.AsyncResultStatus.Failed) {
						if (result.value === Office.MailboxEnums.BodyType.Html) {
							// Body is HTML
							embedString = '<a href="' + uri + '">' + uri + '</a>';
							that.item.body.setSelectedDataAsync(
								embedString,
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

