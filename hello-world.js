var http = require('http');
var server = http.createServer(function(request, response){
    HandleResponse(response);
});

server.listen(8000);

console.log("server running")

function HandleResponse(response) {
    console.log('STATUS: ' + response.statusCode);
    console.log('HEADERS: ' + JSON.stringify(response.headers));
    response.writeHead(200, { "Content-Type": "text/plain" });
    response.end("Hello World new!\n");
}

