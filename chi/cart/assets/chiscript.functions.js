///********************************************************************************************************************************************************************************************************************************************************/
function DuplicateIDCheck(evt) {
    // Warning Duplicate IDs
    var duplicateIDs = new Array();
    $('[id]').each(function () {
        var ids = $('[id="' + this.id + '"]');
        if (ids.length > 1 && ids[0] === this)
            duplicateIDs.push(this.id);
    });
    if (duplicateIDs.length > 0) {
        alert("Multiple elements with same ID found.\r\nPlease fix it, else you may end up with issues when trying to get these elements in jQuery :)\r\n\r\n" + duplicateIDs.toString());
    } else if (evt !== null) {
        alert("No duplicate ids found");
    }
}
///********************************************************************************************************************************************************************************************************************************************************/
function UpdateDebugValues() {
    if (chiConfig.Debug) {
        var $debug_only = $("#debug_only");
        //This method can be invoked multiples times and on each call remove items from the debug area which are not permanent
        //$("#debug_only").find(":not(.permanent)").remove();
        $debug_only.append($(".debug-only"));
        $("#hide_debug_only", $debug_only).on("click", function () {
            $debug_only.hide();
        });
    } else {
        $(".debug-only").remove();
    }
}
///********************************************************************************************************************************************************************************************************************************************************/