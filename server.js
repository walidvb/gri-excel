
const http = require('http'),
    fs = require('fs');

const processPostRequest = require('./processPostRequest')
const excelor = require('./excel')
const port = 7000

const requestHandler = (req, res) => processPostRequest(req, res, () => {
  const project = req.post
  console.log('Received', project)
  if (project.version){
      respondWithXls(project, res)
  }
  else{
    console.log('Project was missing version', project)
    res.writeHead(400)
    res.end()
  }
})

const server = http.createServer((res, ...args) => {
  try{
    requestHandler(res, ...args)
  } catch(err){
    console.log('Error with handler', err)
    res.writeHead(500)
    res.write(JSON.stringify(err))
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
  fs.stat(fileName, function (err, stat) {
    if (err) {
      console.log("not exists: " + fileName);
      res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.write('500 File not found on disk\n');
      res.end();
    }
    console.log('returning file')
    res.writeHead(200, {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      "Content-Disposition": "attachment; filename=" + fileName
    });

    var fileStream = fs.createReadStream(fileName);
    fileStream.pipe(res);
    // res.end()
  }); //end path.exists
}
