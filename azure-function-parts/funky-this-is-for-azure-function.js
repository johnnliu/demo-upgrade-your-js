/*

    this file is not for sharepoint.
    this file is to be run by NodeJS inside Azure Functions

*/
var request = require("request");
var adal = require("adal-node")
var fs = require("fs")

module.exports = function(context, req) {
    context.log('Node.js HTTP trigger function processed a request. RequestUri=%s', req.originalUrl);

    var url = "https://graph.microsoft.com/beta/groups/1b0a3643-4a81-4c39-997e-83c0d0070703/threads/";    
    var authorityHostUrl = 'https://login.microsoftonline.com';
    var tenant = 'johnliu365.onmicrosoft.com';
    var authorityUrl = authorityHostUrl + '/' + tenant;
    var clientId = '37ded58a-5328-41dc-afd1-YOURCLIENTID';
    var clientSecret = 'fEYOURCLIENTSECRETs='
    var resource = 'https://johnliu365.sharepoint.com';
    //var resource = 'https://graph.microsoft.com';

    // for thumbprintand certificate read
    // http://johnliu.net/blog/2016/5/azure-functions-js-and-app-only-updates-to-sharepoint-online
    var thumbprint = '85b8274140YOUR-THUMB-PRINTad70';
    var certificate = fs.readFileSync(__dirname + '/funky.pem', { encoding : 'utf8'});
    
    context.log(certificate);
    
    var authContext = new adal.AuthenticationContext(authorityUrl);

    /*
        call graph with clientId and clientSecret

    authContext.acquireTokenWithClientCredentials(resource, clientId, clientSecret, function(err, tokenResponse) {
        if (err) {
            context.log('well that didn\'t work: ' + err.stack);
            context.done();
        } else {
            context.log(tokenResponse);

            var options = {
                uri: url,
                headers: {
                    'Authorization': 'Bearer ' + accesstoken          
                }
            };
            
            context.log(req.body);
        
            var accesstoken = tokenResponse.accessToken;
            options = { 
                method: 'GET', 
                uri: "https://graph.microsoft.com/beta/groups/1b0a3643-4a81-4c39-997e-83c0d0070703/threads/", 
                headers: { 
                    'Accept': 'application/json;odata.metadata=full',
                    'Authorization': 'Bearer ' + accesstoken
                }
            };
        
            context.log(options);
            request(options, function(error, res, body){
                context.log(error);
                context.log(body);
                context.res = { body: body || '' };
                context.done();
            });

    */

    /*
        call SharePoint Online with client certificate 
    */
    authContext.acquireTokenWithClientCertificate(resource, clientId, certificate, thumbprint, function(err, tokenResponse) {
        if (err) {
            context.log('well that didn\'t work: ' + err.stack);
            context.done();
        } else {
            context.log(tokenResponse);

            var options = {
                uri: url,
                headers: {
                    'Authorization': 'Bearer ' + accesstoken          
                }
            };
            
            context.log(req.body);
        
            var accesstoken = tokenResponse.accessToken;
            options = {
                method: "POST",
                uri: "https://johnliu365.sharepoint.com/_api/web/lists/getbytitle('Poked')/items",
                body: JSON.stringify({ '__metadata': { 'type': 'SP.Data.PokedListItem' }, 'Title': 'Hello, ' + ((req.body ? req.body.name : null) || "test!") }),
                headers: {
                    'Authorization': 'Bearer ' + accesstoken, 
                    'Accept': 'application/json; odata=verbose',
                    'Content-Type': 'application/json; odata=verbose'
                }
            };
            
            context.log(options);
            request(options, function(error, res, body){
                context.log(error);
                context.log(body);
                context.res = { body: body || '' };
                context.done();
            });
        }
    });

    
    

};