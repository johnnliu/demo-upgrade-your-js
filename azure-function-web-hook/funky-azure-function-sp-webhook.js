var request = require("request");
var adal = require("adal-node")
var fs = require("fs")

module.exports = function (context, req) {
    context.log('Node.js HTTP trigger function processed a request. RequestUri=%s', req.originalUrl);

    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = 'johnliu365.onmicrosoft.com';
    var authorityUrl = authorityHostUrl + '/' + tenant;
    var resource = 'https://johnliu365.sharepoint.com';
    var clientId = process.env['MyClientId'];
    var thumbprint = process.env['MyThumbPrint'];
    var certificate = fs.readFileSync(__dirname + '/funky.pem', { encoding: 'utf8' });
    var authContext = new adal.AuthenticationContext(authorityUrl);

    // 1. http://johnliu.net/blog/2016/5/azure-functions-js-and-app-only-updates-to-sharepoint-online 
    // 2. http://johnliu.net/blog/2016/9/working-with-sharepoint-webhooks-with-javascript-using-an-azure-function

    authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function (err, tokenResponse) {
        if (err) {
            context.log('well that didn\'t work: ' + err.stack);
            context.done();
        } else {
            var accesstoken = tokenResponse.accessToken;
            var options = null;
            var headers = {
                'Authorization': 'Bearer ' + accesstoken,
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json'
            };

            if (req.body) {
                context.log(req.body);
            }
            if (req) {
                if (req.query && req.query.validationtoken) {
                    // if validationtoken is specified in query
                    // immediately return token as text/plain
                    context.log(req.query);
                    context.log(req.query.validationtoken);
                    context.res = { "content-type": "text/plain", body: req.query.validationtoken };
                    context.done();
                    return;
                }
                if (req.body && req.body.sub) {
                    // if request with body.sub 
                    // post a subscription request
                    options = {
                        method: 'POST',
                        uri: "https://johnliu365.sharepoint.com/_api/web/lists/getbytitle('subscribe-this')/subscriptions",
                        body: JSON.stringify({
                            "resource": "https://johnliu365.sharepoint.com/_api/web/lists/getbytitle('subscribe-this')",
                            "notificationUrl": "https://johnno-funks.azurewebsites.net/api/poke-spo2?code=q8sq9wxm62asd-YOURTRIGGERCODE",
                            "expirationDateTime": "2017-01-01T16:17:57+00:00",
                            "clientState": "jonnofunks"
                        }),
                        headers: headers
                    };
                    request(options, function (error, res, body) {
                        context.log(error);
                        context.log(body);
                        context.res = { body: body || '' };
                        context.done();
                    });
                    return;
                }
                if (req.body && req.body.subs) {
                    // if request with body.subs 
                    // GET subscriptions request
                    options = {
                        method: 'GET',
                        uri: "https://johnliu365.sharepoint.com/_api/web/lists/getbytitle('subscribe-this')/subscriptions",
                        headers: headers
                    };
                    request(options, function (error, res, body) {
                        context.log(error);
                        context.log(body);
                        context.res = { body: body || '' };
                        context.done();
                    });
                    return;
                }
                // there really should be a DELETE request to remove the subscription
            }

            // default action - create a list item in a different list "Poked"
            options = {
                method: "POST",
                uri: "https://johnliu365.sharepoint.com/_api/web/lists/getbytitle('Poked')/items",
                body: JSON.stringify({ '__metadata': { 'type': 'SP.Data.PokedListItem' }, 'Title': 'Hello, ' + ((req.body ? req.body.name : null) || "test!") }),
                headers: headers
            };

            context.log(options);
            request(options, function (error, res, body) {
                context.log(error);
                context.log(body);
                context.res = { body: body || '' };
                context.done();
            });
        }
    });
};