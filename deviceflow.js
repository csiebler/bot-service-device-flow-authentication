var rp = require('request-promise');
require('dotenv').config();

const AZUREAD_ID = process.env.AZUREAD_ID;
const AZUREAD_APP_ID = process.env.AZUREAD_APP_ID;

function requestDeviceCode(builder, session) {

  console.log("Requesting device code...");

  let options = {
    method: 'POST',
    uri: `https://login.microsoftonline.com/${AZUREAD_ID}/oauth2/devicecode`,
    json: true,
    form: {
      resource: 'https://graph.microsoft.com',
      client_id: AZUREAD_APP_ID
    }
  };

  rp(options)
    .then((body) => {
      console.log(`Body: ${body}`);

      session.privateConversationData.deviceCode = body.device_code;

      let card = new builder.SigninCard(session)
        .text(`Sign-in with code: ${body.user_code}`)
        .button('Sign-in', body.verification_url)
      session.send(new builder.Message(session).addAttachment(card));
    })
    .catch((error) => {
      console.log(`Error: ${error}`);
      session.send(`There was a problem with signing you in. ${error}`);
    });
}

function queryStatus(builder, session) {

  console.log("Waiting until authentication was successful...");
  console.log(`Code: ${session.privateConversationData.deviceCode}`);

  let options = {
    method: 'POST',
    uri: `https://login.microsoftonline.com/${AZUREAD_ID}/oauth2/token`,
    json: true,
    form: {
      'grant_type': 'device_code',
      'resource': 'https://graph.microsoft.com',
      'code': session.privateConversationData.deviceCode,
      'client_id': AZUREAD_APP_ID
    }
  };

  rp(options)
    .then((body) => {
      console.log(`Body: ${body}`);
      session.userData.access_token = body.access_token;
      session.userData.refresh_token = body.refresh_token;
      session.send(`Looks like you've been signed-in!`);
    })
    .catch((error) => {
      console.log(`Error: ${error}`);
      session.send(`There was a problem while querying your authentication status.`);
    });
}

function updateToken(builder, session) {

  let options = {
    method: 'POST',
    uri: `https://login.microsoftonline.com/${AZUREAD_ID}/oauth2/token`,
    json: true,
    form: {
      'scope': 'User.Read',
      'resource': 'https://graph.microsoft.com',
      'refresh_token': session.userData.refresh_token,
      'grant_type': 'refresh_token',
      'client_id': AZUREAD_APP_ID
    }
  };

  rp(options)
    .then((body) => {
      console.log(`Body: ${body}`);
      session.send(`I've updated the access token (HTTP ${response.statusCode})`);
      session.userData.access_token = body.access_token;
    })
    .catch((error) => {
      console.log(`Error: ${error}`);
      session.send(`There was a problem while getting your new access token.`);
    });
}

function getUserInformation(builder, session) {

  let options = {
    method: 'GET',
    uri: `https://graph.microsoft.com/v1.0/me`,
    json: true,
    headers: {
      'Authorization': `Bearer ${session.userData.access_token}`
    }
  }

  rp(options)
    .then((body) => {
      console.log(`Body: ${body}`);
      let userInfo = {
        'id': body.id,
        'firstname': body.givenName,
        'lastname': body.surname,
        'email': body.userPrincipalName
      }
      session.send(JSON.stringify(userInfo));
    })
    .catch((error) => {
      console.log(`Error: ${error}`);
      session.send(`There was a problem while getting your user information.`);
    });
}

module.exports.requestDeviceCode = requestDeviceCode;
module.exports.queryStatus = queryStatus;
module.exports.getUserInformation = getUserInformation;
module.exports.updateToken = updateToken;
