// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

/// <reference path="../../typings/jquery/jquery.d.ts" />
/// <reference path="../../typings/office-js/office-js.d.ts" />
/// <reference path="../../typings/Dropdown.d.ts" />
/// <reference path="home.viewmodel.ts" />
/// <reference path="uri.helper.ts" />

module Controllers {

	export class HomeController {
		vm: ViewModels.HomeViewModel;
		item: Office.Types.ItemCompose;

		constructor() {
			this.vm = new ViewModels.HomeViewModel();
		}

		initializeHandler(reason: Office.InitializationReason) {
			jQuery(document).ready(() => {
				app.initialize();

				if (jQuery.fn.Dropdown) {
					jQuery('.ms-Dropdown').Dropdown();
				}

				jQuery('#clear').click((evt: JQueryEventObject) => { hc.clearForm(evt) });
				jQuery('#submit').click((evt: JQueryEventObject) => { hc.submitForm(evt) });

				hc.item = Office.cast.item.toItemCompose(Office.context.mailbox.item);
				hc.clearForm(null);
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

		bindViewToModel(callback: () => void) {
			var that = this;

			this.vm.title = parseInt(jQuery('#title').val());
			this.vm.name = jQuery('#name').val();
			this.vm.address = jQuery('#address').val();
			this.vm.postcode = jQuery('#postcode').val();
			this.vm.telNumber = jQuery('#telNumber').val();
			this.vm.fromEmailAddress = jQuery('#emailAddress').val();
			this.vm.workSummary = jQuery('#workSummary').val();

			if (this.item.itemType === Office.MailboxEnums.ItemType.Message) {
				var message = Office.cast.item.toMessageCompose(this.item);
				message.to.getAsync(
					(asyncResult: Office.AsyncResult) => {
						that.vm.toEmailAddresses = asyncResult.value
							.map(function(x: Office.EmailAddressDetails) {
								return x.emailAddress;
							});
							callback.apply(this);
					});
			}
			else {
				var appoimtment = Office.cast.item.toAppointmentCompose(this.item);
				appoimtment.requiredAttendees.getAsync(
					(asyncResult: Office.AsyncResult) => {
						that.vm.toEmailAddresses = asyncResult.value
							.map(function(x: Office.EmailAddressDetails) {
								return x.emailAddress;
							});
							callback.apply(this);
					});
			}
		}

		submitForm(evt: JQueryEventObject) {
			this.bindViewToModel(this.setItemBody);
		}

		setItemBody() {
			var that = this;
			var protocol = "est";
			var action = "Action";
			var guid = "1234-1234-1234-1234";
			var uri = helpers.UriBuilder.buildFromObject(protocol, action, guid, this.vm);
			var embedString = "";

			this.item.body.getTypeAsync((asyncResult: Office.AsyncResult) => {
					if (asyncResult.status !== Office.AsyncResultStatus.Failed) {
						if (asyncResult.value === Office.MailboxEnums.BodyType.Html) {
							// Body is HTML
							embedString = '<a href="' + uri + '">' + uri + '</a>';
							that.item.body.setSelectedDataAsync(
								embedString,
								{ coercionType: Office.CoercionType.Html },
								(asyncResult : Office.AsyncResult) => {
									var x = asyncResult;
								}
							)
						} else {
							// Body is text
							that.item.body.setSelectedDataAsync(
								uri,
								{ coercionType: Office.CoercionType.Text },
								(asyncResult: Office.AsyncResult) => { }
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

