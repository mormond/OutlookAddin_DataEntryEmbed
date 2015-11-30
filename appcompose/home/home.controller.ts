// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

/// <reference path="../../typings/jquery/jquery.d.ts" />
/// <reference path="../../typings/office-js/office-js.d.ts" />
/// <reference path="../../typings/Dropdown.d.ts" />
/// <reference path="home.viewmodel.ts" />
/// <reference path="uri.helper.ts" />

module Controllers {

	export class HomeController {
		vm: ViewModels.HomeViewModel;
		messageItem: Office.Types.MessageCompose;
		appointmentItem: Office.Types.AppointmentCompose;

		constructor() {
			this.vm = new ViewModels.HomeViewModel();
		}

		initializeHandler(reason: Office.InitializationReason) {
			var item: Office.Types.ItemCompose;

			jQuery(document).ready(() => {
				app.initialize();

				if (jQuery.fn.Dropdown) {
					jQuery('.ms-Dropdown').Dropdown();
				}

				jQuery('#clear').click((evt: JQueryEventObject) => { hc.clearForm(evt) });
				jQuery('#submit').click((evt: JQueryEventObject) => { hc.submitForm(evt) });

				hc.clearForm(null);
				item = Office.cast.item.toItemCompose(Office.context.mailbox.item);
				hc.messageItem = Office.cast.item.toMessageCompose(item);
				hc.appointmentItem = Office.cast.item.toAppointmentCompose(item);
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

			if (this.messageItem.itemType === Office.MailboxEnums.ItemType.Message) {
				this.messageItem.to.getAsync(
					function(asyncResult: Office.AsyncResult) {
						that.vm.toEmailAddresses = asyncResult.value
							.map(function(x: Office.EmailAddressDetails) {
								return x.emailAddress;
							});
					});
			}
			else {
				this.appointmentItem.requiredAttendees.getAsync(
					function(asyncResult: Office.AsyncResult) {
						that.vm.toEmailAddresses = asyncResult.value
							.map(function(x: Office.EmailAddressDetails) {
								return x.emailAddress;
							});
					});
			}
		}

		submitForm(evt: JQueryEventObject) {
			this.bindViewToModel();
			this.setItemBody();
		}

		setItemBody() {
			var that = this;
			var protocol = "est";
			var guid = "1234-1234-1234-1234";
			var uri = helpers.UriBuilder.buildFromObject(protocol, guid, this.vm);
			var embedString = "";

			this.messageItem.body.getTypeAsync(
				function(asyncResult: Office.AsyncResult) {
					if (asyncResult.status !== Office.AsyncResultStatus.Failed) {
						if (asyncResult.value === Office.MailboxEnums.BodyType.Html) {
							// Body is HTML
							embedString = '<a href="' + uri + '">' + uri + '</a>';
							that.messageItem.body.setSelectedDataAsync(
								embedString,
								{ coercionType: Office.CoercionType.Html },
								function(asyncResult) {
									var x = asyncResult;
								}
							)
						} else {
							// Body is text
							that.messageItem.body.setSelectedDataAsync(
								uri,
								{ coercionType: Office.CoercionType.Text },
								function(asyncResult: Office.AsyncResult) { }
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

