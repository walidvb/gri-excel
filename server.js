
const http = require('http'),
    url = require('url'),
    path = require('path'),
    fs = require('fs');

const processPost = require('./processPost')
const excelor = require('./excel')
const port = 7000

const requestHandler = (req, res) => processPost(req, res, () => {
  const project = req.post
  if (project.version){
      respondWithXls(project, res)
  }
  else{
    res.writeHead(400)
    res.end()
  }
})

const server = http.createServer((...args) => {
  try{
    requestHandler(...args)
  } catch(err){
    res.writeHead(500)
    res.end()
  }
})

server.listen(port, (err) => {
  if (err) {
    return console.log('something bad happened', err)
  }

  console.log(`server is listening on ${port}`)
})


async function respondWithXls(project, res){
  const fileName = await new excelor(project).createDocument()
  fs.exists(fileName, function (exists) {
    if (!exists) {
      console.log("not exists: " + fileName);
      res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.write('500 File not found on disk\n');
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
}
