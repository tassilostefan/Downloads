var fs = require('fs');
var path = require('path');
var http = require('http');

var dir = path.join(__dirname, 'temp');
var source = __filename;
var target = path.join(dir, 'target');
var resP;

console.log("path avail:" +path.dirname);


var server = http.createServer(function(request, response){
    HandleResponse(request, response);
});

server.listen(8080);

console.log("server running");

function HandleResponse(request,response) {
    console.log('ResponseURL: ' + response.url);
    console.log('STATUS: ' + response.statusCode);
    console.log('HEADERS: ' + JSON.stringify(response.headers));
    console.log('RequestURL: ' + request.url);
    //fs.mkdir(dir, handlingError(mkdirped));   
    if(request.url.slice(-3) == "ico"){
        
    } else if(request.url =="/") {
        response.writeHead(200, { "Content-Type": "html" });
        fs.readdir(__dirname, (err, files) => {
 //           response.write('<ol>');
            files.forEach(file => {
                fPath = path.join(__dirname, file);
                response.write('<p><a target=\'blank\' href=\'localhost:8080\\'+file+'\'>'+file+'</a></p>');
                console.log(file);
            });
 //           response.write('</ol>');
            response.end();
          });
    } else {
        fPath = path.join(__dirname, request.url);
        console.log('fPath: ' + fPath);
        resP = response;
        content = fs.readFile(fPath, 'utf-8', handlingError(wroteFile));
//        console.log('content: ' + content);
//        response.write(content);
        
    }
}



function mkdirped(){
    fs.readFile(source, handlingError(wroteFile));
}

function haveFile(content){
    fs.writeFile(target, content, handlingError(wroteFile));
} 

function wroteFile(content){
    resP.writeHead(200, { "Content-Type": "text/plain" });
    resP.write(content);    
    console.log('Done');
    resP.end();    
}

function handlingError(cb){
    return function(err, result) {
        if(err){
            handleError(err);
        }
        else {
            cb(result);
        }
    };
}

function handleError(err){
    console.log("error ocurred: ", err);
}