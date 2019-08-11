'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage() {

    let hostweburl = decodeURIComponent(getQueryStringParameter('SPHostUrl'));
    let appweburl = decodeURIComponent(getQueryStringParameter('SPAppWebUrl'));
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        getUserName();
    });

    // This function prepares, loads, and then executes a SharePoint query to get the current users information
    function getUserName() {
        context.load(user);
        context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
    }

    // This function is executed if the above call is successful
    // It replaces the contents of the 'message' element with the user name
    function onGetUserNameSuccess() {
        $('#message').text('<b>Hello ' + user.get_title() + " Welcome to TND portal</b>");
    }

    // This function is executed if the above call fails
    function onGetUserNameFail(sender, args) {
        alert('Failed to get user name. Error:' + args.get_message());
    }

    // Initialize the RequestExecutor with the app web URL.
    var scriptbase = hostweburl + "/_layouts/15/";
    $.getScript(scriptbase + "SP.RequestExecutor.js", execCrossDomainRequest);

    function execCrossDomainRequest() {
        var executor = new SP.RequestExecutor(appweburl);
    }
    // Button click event.
    $("#btnShowTraining").click(function () {

        var executor = new SP.RequestExecutor(appweburl);
        executor.executeAsync(
            {
                url: appweburl + "/_api/SP.AppContextSite(@target)/web/lists/getbytitle('Training')/Items?@target='"
                    + hostweburl + "'",
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: successHandler,
                error: errorHandler
            }
        );

    });
};

// On query Success
function successHandler(data) {
    var trainingCollection = "";
    var jsonObject = JSON.parse(data.body);
  
    var results = jsonObject.d.results;
    for (var i = 0; i < results.length; i++) {
        trainingCollection = trainingCollection + "<p><h3>" + results[i].Title + "</h3></p><hr>";
    }

    document.getElementById("resultsDiv").innerHTML = trainingCollection;
}
// Error handler.
function errorHandler(error) {
    $("#resultsDiv").append(error.statusText)
}

//Function to get a parameter value by a specific key
function getQueryStringParameter(urlParameterKey) {
    var params = document.URL.split('?')[1].split('&');
    var strParams = '';
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split('=');
        if (singleParam[0] == urlParameterKey)
            return decodeURIComponent(singleParam[1]);
    }
}