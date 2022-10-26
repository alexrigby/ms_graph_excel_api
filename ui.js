// Select DOM elements to work with
const welcomeDiv = document.getElementById("welcomeMessage");
const signInButton = document.getElementById("signIn");
const signOutButton = document.getElementById("signOut");

//gets name from account, adds it to UI welcome message
function showWelcomeMessage(account) {
  welcomeDiv.innerHTML = `Welcome ${account.name}`;
}
