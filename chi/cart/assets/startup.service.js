//http://addyosmani.com/resources/essentialjsdesignpatterns/book/
var StartupService = (function () {
    function StartupService() { }
    var instance;
    var startupItem = [];
    return {
        GetInstance: function () {
            if (instance == null) {
                instance = new StartupService();
                // Hide the constructor so the returned objected can't be new'd...
                instance.constructor = null;
            }
            return instance;
        },
        AddStartupCall: function (startupEntityInstance) {
            try {
                if (typeof startupEntityInstance === 'undefined')
                    throw Error("Invalid parameters for AddStartupCall [startupEntityInstance is undefined]");

                if (typeof startupEntityInstance.Method === "undefined")
                    throw Error("Invalid parameters for AddStartupCall [method is undefined]");

                if (typeof startupEntityInstance.Method !== "function")
                    throw Error("Invalid parameters for AddStartupCall [method is not function]");

                startupItem.push(startupEntityInstance);
            } catch (e) {
                alert(e);
            }
        },
        InvokeStartupCalls: function () {
            for (var i = 0; i < startupItem.length; i++) {
                var startupEntity = startupItem[i];
                startupEntity.Method(startupEntity.Arguments);
            }
        }
    };
})();

function StartupEntity(method, args) {
    try {
        if (typeof method !== "function")
            throw Error("Invalid parameters for StartupEntity [method is not function]");

        this.Arguments = args || {};
        this.Method = method;
    } catch (e) {
        alert(e);
    }
}