var RC = window.RC || {};
RC.GenerateTeamsMeetingLink = function () { };
RC.GenerateTeamsMeetingLink.prototype.getMetadata = function () {
    return {
        boundParameter: null,
        parameterTypes: {},
        operationType: 0, // This is a function. Use '0' for actions and '2' for CRUD
        operationName: "rc_GenerateTeamsMeetingLink",
    };
};

(function () {
    // Code to run in the form OnLoad event
    this.formOnLoad = function (executionContext) {
        var formContext = executionContext.getFormContext();
        formContext.getAttribute("rc_addmicrosoftteamslink").addOnChange(RC.attributeOnChange);
    }

    // Code to run in the attribute OnChange event 
    this.attributeOnChange = async function (executionContext) {
        var formContext = executionContext.getFormContext();

        if(formContext.getAttribute("rc_addmicrosoftteamslink").getValue() == true){
            var generateTeamsMeetingLink = new RC.GenerateTeamsMeetingLink();
            var response = await Xrm.WebApi.online.execute(generateTeamsMeetingLink);
            var responseJson = await response.json();
            
            formContext.getAttribute("description").setValue(responseJson.OnlineMeeting);
        }
        else{
            formContext.getAttribute("description").setValue(null);
        }

    }
}).call(RC);

