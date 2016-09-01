
$(document).on("click", "#poke", function(){
    
    var url = "https://johnno-funks.azurewebsites.net/api/poke-spo?code=i3rmutyxjuxrk4w8umg52vs4icjrkyf8-secret-endpoint";
    var data = {
        "name": $("#poke-by").val()
    };
    
    var promise = $.ajax({
        url: url,
        dataType: "json",
        contentType: 'application/json',
        headers: {"Accept": "application/json; odata=verbose"},
        method: "POST",
        data: JSON.stringify(data)
    });
    
    promise.done(function(response){
    
        var text = JSON.stringify(response.d);
    
        $("#poke-results").text(text);
    });
    
});


