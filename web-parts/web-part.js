
$(document).ready(function(){
    
    var promise = $.ajax({
        url: _spPageContextInfo.webServerRelativeUrl + "_api/web/lists/getByTitle('Documents')/items?$select=FileLeafRef,Title",
        dataType: "json",
        contentType: 'application/json',
        headers: {"Accept": "application/json; odata=verbose"},
        method: "GET",
        cache: false
    });
    
    promise.done(function(response){
    
        // clear existing dom
        $("ul#container").empty();
    
        $.each(response.d.results, function(i, result) {        	
            var $row = $("<li></li>");
            $row.text(result.FileLeafRef);
    
            $("ul#container").append($row);
            // insert each list item
        });
    });
    
});


