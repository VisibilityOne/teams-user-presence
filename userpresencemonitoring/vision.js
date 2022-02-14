var createError = require('http-errors');
var express = require('express');
var app = express();
var path = require('path');
var cookieParser = require('cookie-parser');
var logger = require('morgan');
var https = require('https').Server(app);
var http = require('http').Server(app);
var io = require('socket.io')(http);
const querystring = require('querystring');
const axios = require("axios");
const { isLength } = require("lodash");
const qs = require('querystring');
var delegatedToken = "";
var refresh_token  = "";
var arrDelegatedTokens = [];
var state = "";
const client_id = "b63d093f-daa5-4a8e-9cef-5c4c6f2684c2";
const client_secret = "d7X7Q~P7HP_xiC67hVgWDVAK8G~5jOGPlFtBh";
const morgan = require('morgan');
const { start } = require('repl');
var startPolling = false;
var usersIds = "";
var code ="";

app.use(morgan('tiny'));
app.get('/', async(req, res) => {
  
  if(req.query && req.query.code && req.query.state){
    code = req.query.code;
    state = req.query.state;
    if(code && state)
    {
      console.log("Tenant_id:" + state);
      console.log(code);
    }
    let token = null;
    var maxRetry = 100;
    while(token == null && maxRetry > 0)
    {
      token = await getToken(state, client_id, client_secret);
      maxRetry--;
    }
    console.log(token);

    //let data = await subscribeCallRecords(state);
    //let data = await prodSubscribeCallRecords(state);
    //let data = await cxdetectUpdateUserStatus(state);

    //console.log(data);
    let users = await getUsers(token);
    let value = users.data.value;
    console.log(state);
    console.log(code);
    let ids = [];
    let error = "";
    value.forEach(async (val, index) => {
          ids.push(val.id);
    });
    usersIds = {
      ids: ids
    };

    let outputCode = '';
    let userPresence = "";
    let instruction = "Please do not close this browser or refresh this browser";
    let count = 0;
    try {
      console.log("start getting user presence at request");
      delegatedToken = await getTokenDelegatedVision(state, client_id, client_secret, code);
      
      var tokenDetails = {
        state: state,
        delegatedToken: delegatedToken,
        usersIds: usersIds
      };
      var index = arrDelegatedTokens.findIndex(elem => elem.state == state);
      if(index == -1)
      {
        arrDelegatedTokens.push(tokenDetails);
      }
      else
      {
        arrDelegatedTokens[index] = tokenDetails;
      }

      refresh_token = delegatedToken.refresh_token;
      let userPresenceResult = await getUsersPresence(delegatedToken.access_token, usersIds);
      if(userPresenceResult && userPresenceResult.data)
      {
        userPresence = JSON.stringify(userPresenceResult.data.value);
        let users = userPresenceResult.data.value;
        users.forEach( async(val) => {
          //var updateResult = await updateUserStatus(val);
          await visionUpdateUserStatus(val);
          //await cxdetectUpdateUserStatus(val);
          //console.log(updateResult);
        });
        
        startPolling = true;
      }
    }catch (error) {
      delegatedToken = await getRefreshTokenDelegatedVision(state, client_id, client_secret, refresh_token);
      refresh_token = delegatedToken.refresh_token;
      let userPresenceResultV2 = await getUsersPresence(delegatedToken.access_token, usersIds);
      if(userPresenceResultV2 && userPresenceResultV2.data)
      {
        userPresenceV2 = JSON.stringify(userPresenceResultV2.data.value);
        let users = userPresenceResultV2.data.value;
        users.forEach( async(val) => {
          //var updateResult = await updateUserStatus(val);
          await visionUpdateUserStatus(val);
          //console.log(updateResult);
        });
        startPolling = true;
      }
    }
    console.log("ken");
    //res.send({userPresence});
    code = "";
    state = "";
    res.sendFile(__dirname + '/index.html');

  } else {
    res.json(arrDelegatedTokens);
  }
});

// app.get("/monitoringTeamsActivated", (res, req, next) => {
//   try {
//     var state = req.query.state;
//     var findState = false;
//     findState = arrDelegatedTokens.find( (elem) => {
//         if(elem.state == state) {
//           return true;
//         }
//     });
//     res.json(findState);
//   }catch(error) {

//   }
// });

setInterval(async() => {
  arrDelegatedTokens.forEach( async(details, index) => {
    //await insertUsers(details.state);
    //await prodInsertUsers(details.state);
    //await cxdetectInsertUsers(details.state);
    var delegatedToken = details.delegatedToken;
    var state =  details.state;
    var usersIds = details.usersIds;
    let token = null;
    var maxRetry = 100;
    while(token == null && maxRetry > 0)
    {
      token = await getToken(state, client_id, client_secret);
      maxRetry--;
    }
    console.log(token);
    /* if we can pull new users now then replace the cache users*/
    let userBatch = [];
    let users = await getUsers(token);
    if(users) {
      let value = users.data.value;
      let ids = [];
      let error = "";
  
      value.forEach(async (val, index) => {
            ids.push(val.id);
      });
  
      usersIds = {
        ids: ids
      };
      userBatch.push(ids);
    }

    let link = (users.data['@odata.nextLink'])?users.data['@odata.nextLink']:"";
    while(link)
    {
      let userData = await getNextPageUsers(link,token);
      console.log("Next Page users");
      console.log(userData);
      value = userData.data.value;
      console.log('start insert value: ', value);
      link = (userData.data['@odata.nextLink'])?userData.data['@odata.nextLink']:"";
      let ids = [];
      value.forEach(async (val, index) => {
        ids.push(val.id);
      });
      userBatch.push(ids);
    }
    var refreshToken = delegatedToken.refresh_token;
    userBatch.forEach( async(batch, index) => {
      var usersIds = { ids: batch};
    if(delegatedToken && delegatedToken.access_token && usersIds && startPolling){
      try {
        console.log("start getting user presence");
        let userPresenceResult = await getUsersPresence(delegatedToken.access_token, usersIds);
        console.log(JSON.stringify(userPresenceResult.data.value));
        let users = userPresenceResult.data.value;
        users.forEach( async(val) => {
          //var updateResult = await updateUserStatus(val);
          await visionUpdateUserStatus(val);
          //await cxdetectUpdateUserStatus(val);
        });
  
      }catch(error) {
          try {
        console.log("failed with the access token - retry with refresh token");
            delegatedToken = await getRefreshTokenDelegatedVision(state, client_id, client_secret, refresh_token);
          refresh_token = delegatedToken.refresh_token;
          let userPresenceResultV2 = await getUsersPresence(delegatedToken.access_token, usersIds);
          if(userPresenceResultV2 && userPresenceResultV2.data)
          {
            arrDelegatedTokens[index].delegatedToken = delegatedToken;
            userPresenceV2 = JSON.stringify(userPresenceResultV2.data.value);
            console.log(JSON.stringify(userPresenceResultV2.data.value));
            let users = userPresenceResultV2.data.value;
            users.forEach( async(val) => {
              //var updateResult = await updateUserStatus(val);
              await visionUpdateUserStatus(val);
              //console.log(updateResult);
            });
          }
          else {
            //do we pop it here?
            delete arrDelegatedTokens[index];
          }
        }catch(error) {
          delete arrDelegatedTokens[index];
        }
      }
    }
    });
    console.log("polling set interval");
  });

}, 45000);

// setInterval( async() => {
//   arrDelegatedTokens.forEach( async(details, index) => {
//     var tenant_id = details.state;
// //     //await subscribeCallRecords(tenant_id);
//     await prodSubscribeCallRecords(tenant_id);
// //     //await cxdetectUpdateUserStatus(tenant_id);
//     console.log("Adding webhook notification for call record");
//   });
// }, 3600000);

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

const getTokenDelegated = async (tenant_id, client_id, client_secret, code) => {
  try {
    const config = {
      "grant_type": "authorization_code",
      "scope": "Presence.Read.All Presence.Read",
      "client_id": client_id,
      "client_secret": client_secret,
      "code": code,
      "redirect_uri": "https://visibilityuserpresence.azurewebsites.net/" //testing only
    }
  
    const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
  
    return data;
  } catch(error) {
    return null;
  }
}

const getTokenDelegatedProd = async (tenant_id, client_id, client_secret, code) => {
  try {
    const config = {
      "grant_type": "authorization_code",
      "scope": "Presence.Read.All Presence.Read",
      "client_id": client_id,
      "client_secret": client_secret,
      "code": code,
      "redirect_uri": "https://visibilityuserpresence-prod.azurewebsites.net/"
    }
  
    const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
  
    return data;
  } catch(error) {
    return null;
  }
}

const getTokenDelegatedVision = async (tenant_id, client_id, client_secret, code) => {
  try {
    const config = {
      "grant_type": "authorization_code",
      "scope": "Presence.Read.All Presence.Read",
      "client_id": client_id,
      "client_secret": client_secret,
      "code": code,
      "redirect_uri": "https://visibilityuserpresence-visionpoint.azurewebsites.net/"
    }
  
    const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
  
    return data;
  } catch(error) {
    return null;
  }
}

const getTokenDelegatedVytec = async (tenant_id, client_id, client_secret, code) => {
    try {
      const config = {
        "grant_type": "authorization_code",
        "scope": "Presence.Read.All Presence.Read",
        "client_id": client_id,
        "client_secret": client_secret,
        "code": code,
        "redirect_uri": "https://visibilityuserpresence-vytec.azurewebsites.net/"
      }
    
      const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      });
    
      return data;
    } catch(error) {
      return null;
    }
}

const getTokenDelegatedCxDetect = async (tenant_id, client_id, client_secret, code) => {
  try {
    const config = {
      "grant_type": "authorization_code",
      "scope": "Presence.Read.All Presence.Read",
      "client_id": client_id,
      "client_secret": client_secret,
      "code": code,
      "redirect_uri": "https://visibilityuserpresence-cxdetect.azurewebsites.net/"
    }
    const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
    return data;
  } catch(error) {
    return null;
  }
}
const getToken = async (tenant_id, client_id, client_secret ) => {
  try {
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
  catch(error) {
    return null;
  }
}
async function getNextPageUsers(link, token) {
  try {
    const data = await axios.get(link, {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
    });
    return data;
  }catch(error) {
    return null;
  }
}

async function getUsers(token) {
  try {
    const data = await axios.get('https://graph.microsoft.com/v1.0/users', {
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      }
    });
  
    return data;
  }catch(error) {
    return null;
  }
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
  try {
    const config = {
      "grant_type": "refresh_token",
      "scope": "Presence.Read.All Presence.Read",
      "client_id": client_id,
      "client_secret": client_secret,
      "refresh_token": code,
      "redirect_uri": "https://visibilityuserpresence.azurewebsites.net/" //testing only
    }
  
    const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });
  
    return data;
  }catch(error) {
    return null;
  }

}


  
const getRefreshTokenDelegatedCxdetect = async (tenant_id, client_id, client_secret, code) => {
    try {
      const config = {
        "grant_type": "refresh_token",
        "scope": "Presence.Read.All Presence.Read",
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": code,
        "redirect_uri": "https://visibilityuserpresence-cxdetect.azurewebsites.net/" //testing only
      }
    
      const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      });
    
      return data;
    }catch(error) {
      return null;
    }
  
}
  
  const getRefreshTokenDelegatedProd = async (tenant_id, client_id, client_secret, code) => {
    try {
      const config = {
        "grant_type": "refresh_token",
        "scope": "Presence.Read.All Presence.Read",
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": code,
        "redirect_uri": "https://visibilityuserpresence-prod.azurewebsites.net/" //testing only
      }
    
      const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      });
    
      return data;
    } catch(error) {
      return null;
    }
  }
  const getRefreshTokenDelegatedVision = async (tenant_id, client_id, client_secret, code) => {
    try {
      const config = {
        "grant_type": "refresh_token",
        "scope": "Presence.Read.All Presence.Read",
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": code,
        "redirect_uri": "https://visibilityuserpresence-visionpoint.azurewebsites.net/" //testing only
      }
    
      const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      });
    
      return data;
    } catch(error) {
      return null;
    }
}
  const getRefreshTokenDelegatedVytec = async (tenant_id, client_id, client_secret, code) => {
    try {
      const config = {
        "grant_type": "refresh_token",
        "scope": "Presence.Read.All Presence.Read",
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": code,
        "redirect_uri": "https://visibilityuserpresence-vytec.azurewebsites.net/" //testing only
      }
    
      const { data } = await axios.post(`https://login.microsoftonline.com/${tenant_id}/oauth2/v2.0/token`, qs.stringify(config), {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      });
    
      return data;
    } catch(error) {
      return null;
    }
}  
  
const insertUsers = async(tenant_id) => {
  try {
    const data  = await axios.post(`https://dev-api.visibility.one/teams/insertUsers/${tenant_id}`, JSON.stringify(tenant_id), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return data;
  } catch(error) {
    return null;
  }
}

const cxdetectInsertUsers = async(tenant_id) => {
  try {
    const data  = await axios.post(`https://cxdetect-api.visibility.one/teams/insertUsers/${tenant_id}`, JSON.stringify(tenant_id), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return data;
  } catch(error) {
    return null;
  }
}


const prodInsertUsers = async(tenant_id) => {
  try {
    const data  = await axios.post(`https://api.visibility.one/teams/insertUsers/${tenant_id}`, JSON.stringify(tenant_id), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
    return data;
  } catch(error) {
    return null;
  }
}

const updateUserStatus = async(user) => {
  try {
    const data  = await axios.post(`https://dev-api.visibility.one/teams/updateUserCallStatus`, JSON.stringify(user), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
  
    return data;
  }catch(error) {
    return null;
  }
}

const cxdetectUpdateUserStatus = async(user) => {
  try {
    const data  = await axios.post(`https://cxdetect-api.visibility.one/teams/updateUserCallStatus`, JSON.stringify(user), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
  
    return data;
  }catch(error) {
    return null;
  }
}

const prodUpdateUserStatus = async(user) => {
  try {
    const data  = await axios.post(`https://api.visibility.one/teams/updateUserCallStatus`, JSON.stringify(user), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
  
    return data;
  }catch(error) {
    return null;
  }
}

const visionUpdateUserStatus = async(user) => {
  try {
    const data  = await axios.post(`https://unifyplus-api.visionpointllc.com/teams/updateUserCallStatus`, JSON.stringify(user), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
  
    return data;
  }catch(error) {
    return null;
  }
}

const vytecUpdateUserStatus = async(user) => {
    try {
      const data  = await axios.post(`https://api.vcareanz.com/teams/updateUserCallStatus`, JSON.stringify(user), {
        headers: {
          'Content-Type': 'application/json'
        }
      });
    
      return data;
    }catch(error) {
      return null;
    }
}

const prodSubscribeCallRecords = async(tenant_id) => {
  try {
    const data  = await axios.post(`https://api.visibility.one/teams/subscription/${tenant_id}`, JSON.stringify(tenant_id), {
      headers: {
        'Content-Type': 'application/json'
      }
    });

    return data;
  }catch(error) {
    return null;
  }
};

const cxdetectSubscribeCallRecords = async(tenant_id) => {
  try {
    const data  = await axios.post(`https://cxdetect-api.visibility.one/teams/subscription/${tenant_id}`, JSON.stringify(tenant_id), {
      headers: {
        'Content-Type': 'application/json'
      }
    });

    return data;
  }catch(error) {
    return null;
  }
};

const subscribeCallRecords = async(tenant_id) => {
  try {
    const data  = await axios.post(`https://dev-api.visibility.one/teams/subscription/${tenant_id}`, JSON.stringify(tenant_id), {
      headers: {
        'Content-Type': 'application/json'
      }
    });
  
    return data;  
  } catch(error) {
    return null;
  }
};

module.exports = app;
