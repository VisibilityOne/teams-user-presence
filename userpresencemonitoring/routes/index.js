var express = require('express');
const querystring = require('querystring');
const axios = require("axios");
const { isLength } = require("lodash");
const qs = require('querystring');
var router = express.Router();

/* GET home page. */
router.get('/', async function(req, res, next) {
  let code = req.query.code;
  let state = req.query.state;
  let outputCode = '';
  const client_id = "b63d093f-daa5-4a8e-9cef-5c4c6f2684c2";
  const client_secret = "62mszmxx81_VbjNV-G3s.pHkDjybfW-.BK";
  let userPresence = "";
  let instruction = "Please do not close this browser or refresh this browser";
  let refresh_token = "";
  let delegatedToken;
  let count = 0;
  if(code && state)
  {
    outputCode = code;
    let token = await getToken(state, client_id, client_secret);
    let users = await getUsers(token);
    let value = users.data.value;
    let ids = [];
    let error = "";
    value.forEach(async (val, index) => {
        ids.push(val.id);
    });
    let usersIds = {
        ids: ids
    };
    try {
      delegatedToken = await getTokenDelegated(state, client_id, client_secret, code);
      refresh_token = delegatedToken.refresh_token;
      count++;
        let userPresenceResult = await getUsersPresence(delegatedToken.access_token, usersIds);
        if(userPresenceResult && userPresenceResult.data)
        {
          userPresence = JSON.stringify(userPresenceResult.data.value);
        } else{
          //get new access token using the refresh token
          delegatedToken = await getRefreshTokenDelegated(state, client_id, client_secret, refresh_token);
          refresh_token = delegatedToken.refresh_token;
          userPresenceResult = await getUsersPresence(delegatedToken.access_token, usersIds);
          if(userPresenceResult && userPresenceResult.data)
          {
            userPresence = JSON.stringify(userPresenceResult.data.value);
          }
        }
        console.log(count);
        res.json({ title: 'Microsoft Teams Presence Monitoring', userPresence: userPresence, instruction: instruction, error: error, count:count});
    }catch(error) {
      error = error.message;
      res.render('index', { title: 'Microsoft Teams Presence Monitoring', error: error});
    }
  }

  
});

const getTokenDelegated = async (tenant_id, client_id, client_secret, code) => {
  const config = {
    "grant_type": "authorization_code",
    "scope": "Presence.Read.All Presence.Read",
    "client_id": client_id,
    "client_secret": client_secret,
    "code": code,
    "redirect_uri": "https://teamsuserpresence.azurewebsites.net/" //testing only
  }

  const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  });

  return data;
}

const getToken = async (tenant_id, client_id, client_secret ) => {
  const config = {
    "grant_type": "client_credentials",
    "scope": "https://graph.microsoft.com/.default",
    "client_id": client_id,
    "client_secret": client_secret,
  }

  const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  });

  return data.access_token;
}

async function getUsers(token) {
  const data = await axios.get('https://graph.microsoft.com/v1.0/users', {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${token}`
    }
  });

  return data;
}

async function getUsersPresence(token, users) {
  try {
    let data = await axios.post(`https://graph.microsoft.com/v1.0/communications/getPresencesByUserId`, JSON.stringify(users), {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
    });
    return data;
  } catch (error) {
    return error.message;
  }

}

const getRefreshTokenDelegated = async (tenant_id, client_id, client_secret, code) => {
  const config = {
    "grant_type": "refresh_token",
    "scope": "Presence.Read.All Presence.Read",
    "client_id": client_id,
    "client_secret": client_secret,
    "refresh_token": code,
    "redirect_uri": "https://teamsuserpresence.azurewebsites.net/" //testing only
  }

  const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    }
  });

  return data.access_token;
}


module.exports = router;
