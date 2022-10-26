//MSLA =Microsoft Authentication Library

// const graphConfig = {
//   graphWorksheetsEndpoint:
//     // "https://graph.microsoft.com/v1.0/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets('Data','LinkedActivities')/usedRange",
//     "https://graph.microsoft.com/v1.0/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets('Project Months')/usedRange",
// };

const worksheetsEndpoints = [
  "https://graph.microsoft.com/v1.0/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets('Data')/usedRange",
  "https://graph.microsoft.com/v1.0/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets('Project Months')/usedRange",
  "https://graph.microsoft.com/v1.0/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets('LinkedActivities')/usedRange",
  "https://graph.microsoft.com/v1.0/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets('WP details')/usedRange",
];

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: AUTHORITY_URL,
    redirectUri: REDIRECT_URI,
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set this to "true" if you're having issues on Internet Explorer 11 or Edge
  },
};

// Add scopes for the ID token to be used at Microsoft identity platform endpoints.
const loginRequest = {
  scopes: ["openid", "profile", "User.Read"],
};

// Add scopes for the access token to be used at Microsoft Graph API endpoints.
const tokenRequest = {
  scopes: ["Files.Read"],
};

//constructs userAgen Application object from configs specified in 'authConfig' e.g Client-ID and authority
const myMSALObj = new Msal.UserAgentApplication(msalConfig);

function logWorkbook(data, endpoint) {
  // console.log(data.values); == all values as there type. using text as wrote origional code
  console.log(data.text);
}

//first time user clicks sign in button prompts pop up authentication window
function signIn() {
  myMSALObj
    .loginPopup(loginRequest)
    .then((loginResponse) => {
      if (myMSALObj.getAccount()) {
        showWelcomeMessage(myMSALObj.getAccount());
      }
    })
    .catch((error) => {
      console.log(error);
    });
}

function signOut() {
  myMSALObj.logout();
}

function getTokenPopup(request) {
  //acquiretokensilent should be called if token has been granted already so no need to re-authenticate
  return myMSALObj.acquireTokenSilent(request).catch((error) => {
    console.log(error);
    console.log("silent token acquisition fails. acquiring token using popup");

    // fallback to interaction when the silent call fails
    return (
      myMSALObj
        //acquireTokenPopup opens a popup window to interact with and sign in
        .acquireTokenPopup(request)
        .then((tokenResponse) => {
          return tokenResponse;
        })
        .catch((error) => {
          console.log(error);
        })
    );
  });
}

function getWorkbook() {
  //if there is an account
  if (myMSALObj.getAccount()) {
    for (let i = 0; i < worksheetsEndpoints.length; i++) {
      //get the token
      getTokenPopup(tokenRequest)
        //then with the tokens response
        .then((response) => {
          // console.log(response);
          //make an API call to MS Graph
          callMSGraph(worksheetsEndpoints[i], response.accessToken, logWorkbook);
        })
        .catch((error) => {
          console.log(error);
        });
    }
  }
}

function callMSGraph(endpoint, token, callback) {
  const headers = new Headers();
  //bearer is the token with 'Bearer' before it
  const bearer = `Bearer ${token}`;

  headers.append("Authorization", bearer);
  const options = {
    method: "GET",
    headers: headers,
  };

  fetch(endpoint, options)
    .then((response) => response.json())
    .then((response) => callback(response, endpoint))
    .catch((error) => console.log(error));
}

// function seeProfile() {
//   if (myMSALObj.getAccount()) {
//     getTokenPopup(loginRequest)
//       .then((response) => {
//         callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
//         profileButton.classList.add("d-none");
//       })
//       .catch((error) => {
//         console.log(error);
//       });
//   }
// }

// function callMSGraph(theUrl, accessToken, callback) {
//   var xmlHttp = new XMLHttpRequest();
//   xmlHttp.onreadystatechange = function () {
//     if (this.readyState == 4 && this.status == 200) {
//       callback(JSON.parse(this.responseText));
//     }
//   };
//   xmlHttp.open("GET", theUrl, true); // true for asynchronous
//   xmlHttp.setRequestHeader("Authorization", "Bearer " + accessToken);
//   xmlHttp.send();
// }

// const options = {
//   authProvider,
// };

// const client = Client.init(options);

// let worksheets = client.api("/me/drive/items/01CGANUMA2V6VFYNHFLZGIU2GD5Y2UGMJ3/workbook/worksheets").get();

// console.log(worksheets);
