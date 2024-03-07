function getRates(currency) {
  let api_key = '';
  let url = "http://api.exchangeratesapi.io/v1/latest?access_key=" + api_key + "&currencies=" + currency;

  let response = UrlFetchApp.fetch(url);
  return JSON.parse(response.getContentText());
}