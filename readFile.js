var path = require('path');
var http = require('http');
var fs = require('fs');
var childProcess = require('child_process');
var resP;

var listenPort = 'sudo fuser -k 5050/tcp';

/* childProcess.exec('sudo fuser -k 5050/tcp', (err, stdout, stderr) => {
    if (err) {
      // node couldn't execute the command
      console.log(`err: ${err}`);
      return;
    }
  
    // the *entire* stdout and stderr (buffered)
    console.log(`stdout: ${stdout}`);
    console.log(`stderr: ${stderr}`);
  });
 */


//2.
var server = http.createServer(function (request, response) {
    var isDir = false;
    var isIco = request.url.slice(-3) == "ico";
    var isJs = request.url.slice(-3) == ".js";
    var dPath = path.join(__dirname, request.url);
    if(!isIco){
        isDir = fs.lstatSync(dPath).isDirectory();
    }
    console.log('ResponseURL: ' + response.url);
    console.log('STATUS: ' + response.statusCode);
    console.log('HEADERS: ' + JSON.stringify(response.headers));
    console.log('RequestURL: ' + request.url);
    console.log('dPath: ' + dPath + " isDir: " + isDir );
    //fs.mkdir(dir, handlingError(mkdirped));   
    if(isIco){
        console.log('request for ico, ending request...');
        response.end();       
    } else if(isJs) {
        runScript(dPath, function (error, pgResp) {
            if (error) throw error;
            console.log('finished running ' + dPath);
            if (error) {
                response.writeHead(404);
                response.write('Contents you are looking are Not Found');
            } else {
                response.writeHead(200, {
                    'Content-Type': 'text/plain'
                });
                response.write(pgResp);
            }
            response.end();
        });
    } else if(isDir) {
        response.writeHead(200, { "Content-Type": "html" });
        fs.readdir(dPath, (err, files) => {
 //           response.write('<ol>');
            files.forEach(file => {
                fPath = path.join(__dirname, file);
                response.write('<p><a target=\'blank\' href=\'http://192.168.2.164:5050\\'+file+'\'>'+file+'</a></p>');
                console.log(file);
            });
 //           response.write('</ol>');
            response.end();
          });
    } else {
        fPath = dPath;
        console.log('fPath: ' + fPath);
        resP = response;
        content = fs.readFile(fPath, 'utf-8', handlingError(wroteFile));
//        console.log('content: ' + content);
//        response.write(content);
        
    }

 });
//5.
server.listen(5050);
 
console.log('Server Started listening on 5050');

function runScript(scriptPath, callback) {

    // keep track of whether callback has been invoked to prevent multiple invocations
    var invoked = false;

    var process = childProcess.fork(scriptPath);

    // listen for errors as they may prevent the exit event from firing
    process.on('error', function (error) {
        if (invoked) return;
        invoked = true;
        callback(error);
    });

    // execute the callback once the process has finished running
    process.on('exit', function (code) {
        if (invoked) return;
        invoked = true;
        var error = code === 0 ? null : new Error('exit code ' + code);
        callback(error);
    });

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
    console.log('File content written to response');
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