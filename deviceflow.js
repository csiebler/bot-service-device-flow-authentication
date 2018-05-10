var request = require('request');
require('dotenv').config();

const AZUREAD_ID = process.env.AZUREAD_ID;
const AZUREAD_APP_ID = process.env.AZUREAD_APP_ID;

var headers = {
  'Content-Type': 'application/x-www-form-urlencoded'
}

function requestDeviceCode(builder, session) {

  console.log("Requesting device code...");

  let options = {
    url: `https://login.microsoftonline.com/${AZUREAD_ID}/oauth2/devicecode`,
    method: 'POST',
    headers: headers,
    json: true,
    form: { 'resource': 'https://graph.microsoft.com',
            'client_id': AZUREAD_APP_ID }
  }

  request(options, function (error, response, body) {

    console.log(`(HTTP ${response.statusCode}) Body: ${JSON.stringify(body)}`);

    if (!error && response.statusCode == 200) {
      var code = body.user_code;
      var deviceCode = body.device_code;
      var url = body.verification_url;
      session.privateConversationData.code = code;
      session.privateConversationData.deviceCode = deviceCode;
      session.privateConversationData.url = url;

      let card = new builder.SigninCard(session)
        .text(`Sign-in with device code: ${code}`)
        .button('Sign-in', url)
      session.send(new builder.Message(session).addAttachment(card));
    } else {
      console.log(`Error: ${error}`);
      session.send('There was a problem with signing you in.');
    }
  })

}

function queryStatus(builder, session) {
  
  console.log("Waiting until authentication was successful...");
  console.log(`Code: ${session.privateConversationData.deviceCode}`);

  let options = {
    url: `https://login.microsoftonline.com/${AZUREAD_ID}/oauth2/token`,
    method: 'POST',
    headers: headers,
    json: true,
    form: {
      'grant_type': 'device_code',
      'resource': 'https://graph.microsoft.com',
      'code': session.privateConversationData.deviceCode,
      'client_id': AZUREAD_APP_ID
    }
  }

  request(options, function (error, response, body) {
    console.log(`(HTTP ${response.statusCode}) Body: ${JSON.stringify(body)}`);
    session.userData.access_token =  body.access_token;
    session.userData.refresh_token =  body.refresh_token;
    session.send(`Looks like you've been signed-in!`);
  })
}

function updateToken(builder, session) {
  let options = {
    url: `https://login.microsoftonline.com/${AZUREAD_ID}/oauth2/token`,
    method: 'POST',
    headers: headers,
    json: true,
    form: {
      'scope': 'User.Read',
      'resource': 'https://graph.microsoft.com',
      'refresh_token': session.userData.refresh_token,
      'grant_type': 'refresh_token',
      'client_id': AZUREAD_APP_ID
    }
  }
  request(options, function (error, response, body) {
    console.log(`(HTTP ${response.statusCode}) Body: ${JSON.stringify(body)}`);
    session.send(`I've updated the access token (HTTP ${response.statusCode})`);
    session.userData.access_token =  body.access_token;
  })
}

function getUserInformation(builder, session) {
  let options = {
    url: `https://graph.microsoft.com/v1.0/me`,
    method: 'GET',
    json: true,
    headers: {
      'Authorization': `Bearer ${session.userData.access_token}`
    }
  }

  request(options, function (error, response, body) {
    console.log(`(HTTP ${response.statusCode}) Body: ${JSON.stringify(body)}`);
    let userInfo = {
      'id': body.id,
      'firstname': body.givenName,
      'lastname': body.surname,
      'email': body.userPrincipalName
    }
    session.send(JSON.stringify(userInfo));
    return userInfo;
  })
}

module.exports.requestDeviceCode = requestDeviceCode;
module.exports.queryStatus = queryStatus;
module.exports.getUserInformation = getUserInformation;
module.exports.updateToken = updateToken;
