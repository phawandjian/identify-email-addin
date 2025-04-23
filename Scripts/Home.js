/* Home.js */

(function () {
    "use strict";

    // Hardcoded list of verified senders
    var verifiedSenders = [
        "support@bizkeyhub.com",
        "ariana@bizkeyhub.com",
        "hawandj@gmail.com"
    ];

    // When Office is ready, we wire up item change events
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Outlook) {
            // Add a handler so that whenever the selected item changes, we update the badge
            Office.context.mailbox.addHandlerAsync(
                Office.EventType.ItemChanged,
                onItemChanged
            );

            // Also update badge right away for the current item
            onItemChanged();
        }
    });

    // This function runs every time the user opens or changes the selected item
    function onItemChanged() {
        // Make sure we have a message item
        var item = Office.context.mailbox.item;
        if (item && item.itemType === Office.MailboxEnums.ItemType.Message) {
            // Get the sender (From) property
            var fromEmail = item.from ? (item.from.emailAddress || "").toLowerCase() : "";

            // Check if it's in our verified list
            var isVerified = verifiedSenders.indexOf(fromEmail) !== -1;

            // Update the UI
            var container = document.getElementById("badgeContainer");
            if (container) {
                if (isVerified) {
                    container.innerHTML = "<div class='badge badge-verified'>Verified Sender: " + fromEmail + "</div>";
                } else {
                    container.innerHTML = "<div class='badge badge-unverified'>Not Verified: " + fromEmail + "</div>";
                }
            }
        }
    }

})();
