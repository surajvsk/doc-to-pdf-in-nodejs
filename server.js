const express =  require('express');
const app =  express();
const port = 5000;
const { exec } = require('child_process');
const { v4: uuidv4 } = require('uuid');
const fileUpload = require('express-fileupload');
const fs = require('fs');
app.use(express.json({
    limit: "200mb"
  }));
  app.use(
    express.urlencoded({
      extended: true,
      limit: "200mb",
    })
  );


app.use(fileUpload());

app.post('/', function(req, res,next){
    console.log('req', req.files.inputfile)

    let uploaded_path = __dirname+'/doc/' + req.files.inputfile.name;
    fs.writeFile(uploaded_path, req.files.inputfile.data, function(err) {
        if(err) {
            return console.log(err);
        }
        ///console.log("The file was saved!");
        const cmd = `$documents_path = "${uploaded_path}"
        $output_path = '${__dirname}/doc/${uuidv4()}.pdf'
        $word_app = New-Object -ComObject Word.Application
        # This filter will find .doc as well as .docx documents
        Get-ChildItem -Path $documents_path -Filter *.doc? | ForEach-Object {
            $document = $word_app.Documents.Open($_.FullName)
            $pdf_filename = "$output_path"
            $document.SaveAs([ref] $pdf_filename, [ref] 17)
            $document.Close()
        }
        $word_app.Quit()
        return $output_path`; 
        exec(cmd, {'shell':'powershell.exe'}, (error, stdout, stderr)=> {
            console.log('Error', stderr)
               console.log('OUTPUT', stdout);
               res.status(200).json({status:200, msg:"success", stdout: stdout})
        })
    }); 

})


app.listen(port, function(){
    console.log('Server is running on port', port);
})