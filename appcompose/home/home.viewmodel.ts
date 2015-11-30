// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

module ViewModels {

	export enum Title {
		"Mr",
		"Miss",
		"Mrs",
		"Ms"
	}

	export class HomeViewModel {
		title: Title;
		name: string;
		address: string;
		postcode: string;
		telNumber: string;
		fromEmailAddress: string;
		toEmailAddresses: string[];
		workSummary: string;

		constructor() {
			this.initialise;
		}

		initialise() {
			this.title = Title.Mr;
			this.name = "";
			this.address = "";
			this.postcode = "";
			this.telNumber = "";
			this.fromEmailAddress = "";
			this.toEmailAddresses = [];
			this.workSummary = "";
		}
	}
}
