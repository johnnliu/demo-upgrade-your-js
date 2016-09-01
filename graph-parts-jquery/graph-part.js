
$(document).ready(function(){

    //authorization context 
    var resource = 'https://graph.microsoft.com'; 
    //'https://johnliu365.sharepoint.com'; 
    //var endpoint = 'https://johnliu365.sharepoint.com/_api/web'; 
    var authContext = new AuthenticationContext({ 
        instance: 'https://login.microsoftonline.com/', 
        tenant: 'common', 
        clientId: '5f01318a-a726-4953-b023-56fac693abe8', 
        postLogoutRedirectUri: window.location.origin, 
        cacheLocation: 'localStorage' 
    });
    
    authContext.handleWindowCallback();
    
    var user = authContext.getCachedUser();
    
    if (!user) {
        var $login = $("<input type='button' value='login'/>");
        
        $login.click(function(){
            authContext.login();
        });
        
        $("#user").append($login);
        return;
    }

    $("#user").text(user.userName);

    authContext.acquireToken(resource, function (error, token) {
        // token or die
        
        if (error) {
            $("#token").text(error);
            return;        
        }

        $("#token").text(token);
        
        var promise = $.ajax({ 
            url: "https://graph.microsoft.com/beta/groups/1b0a3643-4a81-4c39-997e-83c0d0070703/threads/", 
            headers: { 
                'Accept': 'application/json;odata.metadata=full',
                'Authorization': 'Bearer ' + token
            },
            method: 'GET', 
            cache: false
        });
        
        promise.done(function (response) { 
            
            $("#container").empty();
            $.each(response.value, function(i, v){
                var $li = $("<li class='thread' ><span class='topic' /><ul class='posts' /></li>");
                $li.find('span.topic').text(v.topic);
                $li.attr("nav", v["posts@odata.navigationLink"]);                
                $("#container").append($li);
            });
        });                        

        $(document).on("click", "li.thread", function(){
            
            var $li = $(this);
            var url = $li.attr("nav");
            var $posts = $li.find("ul.posts");
            
            $.ajax({ 
                method: 'GET', 
                url: url, 
                headers: { 
                    'Accept': 'application/json;odata.metadata=full',
                    'Authorization': 'Bearer ' + token
                }, 
            }).success(function (data) { 
                
                $posts.empty();
                $.each(data.value, function(i, v){
                    var $post = $("<li class='post' ></li>");
                    var $content = $(v.body.content);
                    $post.append($content[0]);
                    $posts.append($post);
                });
            });                        
            
            
        });
            
    });


    authContext.acquireToken("https://johnliu365.sharepoint.com", function(error, token){
    
        $("#token2").text(token);
    });

    
});


