
Office.onReady(function () {
    document.getElementById("injectButton").onclick = function () {
        var html = document.getElementById("htmlContent").value;
        Office.context.mailbox.item.body.setAsync(
            html,
            { coercionType: Office.CoercionType.Html },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    alert("HTML Injected Successfully!");
                } else {
                    alert("Failed: " + asyncResult.error.message);
                }
            }
        );
    };
});
