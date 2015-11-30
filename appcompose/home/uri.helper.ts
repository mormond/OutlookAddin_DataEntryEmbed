// Copyright (c) Microsoft 2015. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

module helpers {
	
	class KeyValuePair {
		key: string;
		value: string;
	}

	export class UriBuilder {

		static buildFromObject(protocol: string, guid: string, obj: any) : string {

			var keyValPairs = UriBuilder.propsToKeyValuePairs(obj);
			var paramString = "";
			
			for (var index = 0; index < keyValPairs.length; index++) {
				var element = keyValPairs[index];

				if (index === 0) {
					paramString += encodeURI("?" + element.key + "=" + element.value);
				}
				else {
					paramString += encodeURI("&" + element.key + "=" + element.value);
				}
			}

			var Uri = protocol +
				"://Action/" +
				guid +
				paramString;
				
				return Uri;
		}

		static propsToKeyValuePairs(obj: any) : KeyValuePair[]{

			return Object.keys(obj)
				.map(function(x) {
					var rObj = new KeyValuePair();
					rObj.key = x;
					rObj.value = obj[x];
					return rObj;
                });
		}


	}



}