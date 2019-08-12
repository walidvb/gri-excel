
const http = require('http'),
    url = require('url'),
    path = require('path'),
    fs = require('fs');

const processPost = require('./processPost')
const excelor = require('./excel')
const port = 3000

const requestHandler = (req, res) => processPost(req, res, async () => {
  console.log(req.post)
  const fileName = await new excelor(req.post).createDocument()
  fs.exists(fileName, function (exists) {
    if (!exists) {
      console.log("not exists: " + fileName);
      res.writeHead(200, { 'Content-Type': 'text/plain' });
      res.write('404 Not Found\n');
      res.end();
    }
    console.log('returning file')
    res.writeHead(200, { 
      'Content-Type': 'application/vnd.ms-excel', 
      "Content-Disposition": "attachment; filename=" + fileName
    });

    var fileStream = fs.createReadStream(fileName);
    fileStream.pipe(res);
    // res.end()
  }); //end path.exists
})

const server = http.createServer(requestHandler)

server.listen(port, (err) => {
  if (err) {
    return console.log('something bad happened', err)
  }

  console.log(`server is listening on ${port}`)
})