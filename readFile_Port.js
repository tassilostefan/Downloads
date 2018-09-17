var path = require('path');
var http = require('http');
var fs = require('fs');
var childProcess = require('child_process');
var resP;
var port = 5050;
var isPortFree = false;

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
isPortTaken(port, (err) => {
    console.log('port ' + port + ' err: ' + err);
    isPortFree = !err;
    if(!err) {
        createServer();
    } 
    
  });

  function isPortTaken(port, fn) {
    var net = require('net')
    var tester = net.createServer()
      .once('error', function(err) {
          if(port < 10000){
              port++;
              console.log('new port no.: ' + port);
              isPortTaken(port, fn);
          } else {
            return fn(err);
        }
      })
      .once('listening', function() {
        tester.once('close', function() {
            fn(false);
        }).close()
      })
      .listen(port);
  };

  console.log('port at ' + port + ' is free: ' + isPortFree);

function createServer(){

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
        if(isIco){
            console.log('request for ico, ending request...');
            response.end();       
        } else if(isJs) {
            runScript(dPath, function (error, pgResp) {
                console.log('finished running ' + dPath);
                if (error) {
                    response.writeHead(404);
                    response.write('Contents you are looking are Not Found/n');
                    response.write('Error: ' + error);
                } else if (typeof pgResp !== 'undefined' && pgResp !== null){
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
                files.forEach(file => {
                    fPath = path.join(__dirname, file);
                    response.write('<p><a target=\'blank\' href=\'http://192.168.2.164:5050\\'+file+'\'>'+file+'</a></p>');
                    console.log(file);
                });
                response.end();
              });
        } else {
            fPath = dPath;
            console.log('fPath: ' + fPath);
            resP = response;
            content = fs.readFile(fPath, 'utf-8', handlingError(wroteFile));
            
        }
    
     });
    server.listen(port);
    server.on('error', function (e) {
        if (e.code == 'EADDRINUSE') {
          port++;
          console.log('Port in use, retrying...new port: ' + port);
          setTimeout(function () {
            server.close();
            server.listen(port);
          }, 1000);
        }
      })
     
    console.log('Server Started listening on' + port);
}




  
  /* Usage */
  
  

function isPortAvailable(port, fn) {
    var net = require('net');
    var tester = net.createServer()
    .once('error', function (error) {
      //if (error.code != 'EADDRINUSE') return fn(error);
      console.log('port ' + port + ' is free: ..false..'  );
      console.log('error.code: '+ error.code);
      fn(null, false);
    })
    .once('listening', function() {
      tester.once('close', function() { fn(null, true); })
      .close();
      console.log('listening on port ' + port);
    })
    .listen(port);
    console.log('listen port ' + port);
}

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