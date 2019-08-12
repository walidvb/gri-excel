var http = require('http');
var querystring = require('querystring');

function processPost(request, response, callback) {
    var queryData = "";
    if(typeof callback !== 'function') return null;

  if (request.method === 'POST' || request.method == 'PUT') {
        request.on('data', function(data) {
            queryData += data;
            if(queryData.length > 1e6) {
                queryData = "";
                response.writeHead(413, {'Content-Type': 'text/plain'}).end();
                request.connection.destroy();
            }
        });

        request.on('end', function() {
          request.post = JSON.parse(queryData);
            callback();
        });

    } else {
      callback(request, response)
      response.writeHead(405, {'Content-Type': 'text/plain'});
      response.end();
    }
}
module.exports = processPost