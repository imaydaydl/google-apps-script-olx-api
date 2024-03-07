var olx_url = {'pt': 'https://www.olx.pt'};
var olx_client_id = {'pt': ''};
var olx_client_secret = {'pt': ''};
var olx_version = '2.0';
var client_token = {'pt': ''};
var olx_token = {'pt': ''};

var clientCredentials = function(c = 'pt') {
  let urlencoded = {
    "grant_type": "client_credentials",
    "scope": "v2 read write",
    "client_id": olx_client_id[c],
    "client_secret": olx_client_secret[c]
  };

  let requestOptions = {
    'method': 'POST',
    'contentType': "application/json",
    'payload': JSON.stringify(urlencoded)
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/open/oauth/token", requestOptions);
  let credentials = JSON.parse(query.getContentText());
  client_token[c] = credentials['access_token'];
}

var authorizationCode = function(olx_code, c = 'pt') {
  let urlencoded = {
    "grant_type": "authorization_code",
    "scope": "v2 read write",
    "client_id": olx_client_id[c],
    "client_secret": olx_client_secret[c],
    "code": olx_code
  };

  let requestOptions = {
    'method': 'POST',
    'contentType': "application/json",
    'payload': JSON.stringify(urlencoded)
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/open/oauth/token", requestOptions);
  let code = JSON.parse(query.getContentText());
  console.log(code);
  olx_token[c] = code['access_token'];

  return code;
}

var refreshToken = function(c = 'pt') {
  let urlencoded = {
    "grant_type": "refresh_token",
    "client_id": olx_client_id[c],
    "client_secret": olx_client_secret[c],
    "refresh_token": ''
  };

  let requestOptions = {
    'method': 'POST',
    'contentType': "application/json",
    'payload': JSON.stringify(urlencoded)
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/open/oauth/token", requestOptions);
  let refresh = JSON.parse(query.getContentText());
}

var getUserAdverts = function(c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'GET'
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/adverts", requestOptions);
  console.log(query.getContentText());
  return JSON.parse(query.getContentText());
}

var createAdvert = function(data, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'POST',
    'contentType': "application/json",
    'payload': JSON.stringify(data),
    "muteHttpExceptions": true
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/adverts", requestOptions);
  return JSON.parse(query.getContentText());
}

var updateAdvert = function(id, data, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'PUT',
    'contentType': "application/json",
    'payload': JSON.stringify(data)
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/adverts/" + id, requestOptions);
  return JSON.parse(query.getContentText());
}

var deleteAdvert = function(id, c = 'pt', data = '') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'DELETE',
    'payload': data
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/adverts/" + id, requestOptions);
  return JSON.parse(query.getContentText());
}

var getAdvertStats = function(id, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'GET'
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/adverts/" + id + "/statistics", requestOptions);
  return JSON.parse(query.getContentText());
}

var getCategories = function(c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + client_token[c], "Version": olx_version},
    'method': 'GET'
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/categories", requestOptions);
  return JSON.parse(query.getContentText());
}

var getCities = function(data, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + client_token[c], "Version": olx_version},
    'method': 'GET',
    'muteHttpExceptions': true
  };

  let url = olx_url[c] + "/api/partner/cities?offset=" + data.offset;

  let query = UrlFetchApp.fetch(url, requestOptions);
  console.log(1, query.getContentText());
  return JSON.parse(query.getContentText());
}

var getThreads = function(c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'GET'
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/threads", requestOptions);
  return JSON.parse(query.getContentText());
}

var getMessages = function(thread_id, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'GET'
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/threads/" + thread_id + "/messages", requestOptions);
  return JSON.parse(query.getContentText());
}

var sendMessages = function(thread_id, text, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'POST',
    'contentType': "application/json",
    'payload': JSON.stringify({"text": text})
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/threads/" + thread_id + "/messages", requestOptions);
  return JSON.parse(query.getContentText());
}

var getUser = function(user_id, c = 'pt') {
  let requestOptions = {
    'headers': {"Authorization": "Bearer " + olx_token[c], "Version": olx_version},
    'method': 'GET'
  };

  let query = UrlFetchApp.fetch(olx_url[c] + "/api/partner/users/" + user_id, requestOptions);
  return JSON.parse(query.getContentText());
}