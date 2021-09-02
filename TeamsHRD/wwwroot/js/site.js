// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

//this function will attempt to sign-in using Teams. It calls the getTeamsToken function in this file
signInWithTeams = async () => {
	try {

		const result = await this.getTeamsToken();
		//$("#consentButton").show();
		showSharepointButtons();
		return

		//todo
		//return await this.handleMsalToken(result);
	}
	catch (error) {
		console.log("Teams SSO Error: " + error);
		if (typeof error === "string" && error.toLowerCase() == "resourcedisabled") {
			console.log("Resource Disabled - probably something wrong with your manifest or Azure AD App Reg config");
			$("#consentButton").show();
			return
		}

		else {
			console.warn("error... but, it wasn't resource disabled....")
			$("#consentButton").show();
			return
        }
	}
}


//called on every page loaded in the client app - this is where it all starts for authentication in the client app
//this function will calls the initialiseTeams function and reports back if the page has been opened in Teams or not
//then it sets the variable inTeams to true to false depending on what is returned by the initaliseTeams function
//if page is loaded in Teams, then it checks if the config page is being accessed and sets the options appropriately to handle this
//if it isn't the config page then signinWithTeams function is called, that handles SSO from within Teams

handlePageLoad = async () => {

	try {
		await this.initialiseTeams();
		this.inTeams = true;
	}
	catch {
		this.inTeams = false;
	}

	if (this.inTeams) {

		//Check for config page...
		if (window.location.pathname == "/config") {
			microsoftTeams.settings.registerOnSaveHandler((saveEvent) => {
				microsoftTeams.settings.setSettings({
					//todo: move this to config
					websiteUrl: "https://configpageurl.com",
					contentUrl: "https://configpageurl.com",
					entityId: "sampleApp",
					suggestedDisplayName: "My New Suggested Tab Name"
				});
				saveEvent.notifySuccess();
			});
			microsoftTeams.settings.setValidityState(true);
		}

		if (await this.signInWithTeams()) {

		}
		else {

			//this is where you can add logic/functions to handle SSO if the page is loaded outside of Teams. Not required for this sample...

			return;
		}
	}
}


//this function initialises the Teams SDK, this is required to pass context and allow the SSO flow to be used in Teams
//success/failure is returned to the function that called this


initialiseTeams = () => {
	console.log("inside initialiseTeams");
	let rejectPromise = null;
	let timeout = null;

	const promise = new Promise((resolve, reject) => {
		rejectPromise = reject;
		microsoftTeams.initialize(() => {
			window.clearTimeout(timeout);
			resolve(true);
		}
		)
	});

	//todo: try and improve this if possible. At present this function will return that the page is not open in Teams
	//based on the MicrosoftTeams.initialize function timing out within 2 seconds

	timeout = window.setTimeout(() => {
		rejectPromise("Teams Initialise Timeout");
	}, 2000);

	return promise;
}

//this function is used to get the Azure AD Access token using the new Teams SSO flow
//uses the settings from within the Teams manifest to perform this SSO

getTeamsToken = () => {

	return new Promise((resolve, reject) => {
		microsoftTeams.authentication.getAuthToken({
			successCallback: (token) => {
				resolve(token);
			},
			failureCallback: (reason) => {
				reject(reason);
			}
		})
	});
}

//used for grabbing the token info after a fallback auth (for testing purposes)
function printAccessToken(result) {
	console.log("access token " + result.accessToken);
	console.log("id token " + result.idToken);
	console.log("type" + result.tokenType);
}

//displays the buttons for redirecting to the appropriate SharePoint site
function showSharepointButtons() {
	$("#consentButton").hide();
	//todo: only show 1 button, and dynamically generate the SharePoint link URL based on Graph API call
	$("#sharepointButton1").show();
	$("#sharepointButton2").show();
}


//this function is the fallback method for sign-in. It uses the Teams SDK to open the /auth-start.html page in a pop-up window.
//This is required, as Azure AD (and many other auth providers) do not support iFraming. microsoftTeams.authentication.authenticate opens a pop-up, not an iFrame.
function fallbackAuth() {

	microsoftTeams.authentication.authenticate({
		url: window.location.origin + "/auth-start.html",
		width: 600,
		height: 535,
		successCallback: function (result) {
			//printAccessToken(result);
			showSharepointButtons();
		},
		failureCallback: function (reason) {
			handleAuthError(reason);
		}
	});
}

function sharepointRedirect1() {
	window.location.assign('https://m365x265372.sharepoint.com/_layouts/15/sharepoint.aspx?');
}

function sharepointRedirect2() {
	window.location.assign('https://en.wikipedia.org/wiki/Microsoft');
}


$(this.handlePageLoad);