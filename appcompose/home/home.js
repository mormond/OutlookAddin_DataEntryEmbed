(function () {
  'use strict';

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    jQuery(document).ready(function () {
      app.initialize();

      if ($.fn.Dropdown) {
        $('.ms-Dropdown').Dropdown();
      }

jQuery('#clear').click(clearForm);
jQuery('#submit').click(submitForm);

    });
  };
  
  function clearForm() {
    
  }
  
  function submitForm() {
    
  }

  function setSubject() {
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.setAsync('Hello world!');
  }

  function getSubject() {
    Office.cast.item.toItemCompose(Office.context.mailbox.item).subject.getAsync(function (result) {
      app.showNotification('The current subject is', result.value);
    });
  }

  function addToRecipients() {
    var item = Office.context.mailbox.item;
    var addressToAdd = {
      displayName: Office.context.mailbox.userProfile.displayName,
      emailAddress: Office.context.mailbox.userProfile.emailAddress
    };

    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      Office.cast.item.toMessageCompose(item).to.addAsync([addressToAdd]);
    } else if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
      Office.cast.item.toAppointmentCompose(item).requiredAttendees.addAsync([addressToAdd]);
    }
  }

})();
