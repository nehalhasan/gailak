
const express = require('express');
const nodemailer= require ('nodemailer');

const path = require('path');
const excel = require('exceljs');
const bodyParser = require('body-parser');

//var wb = new xl.Workbook();
var app = express();
//var router=express();
const port=process.env.port || 3003
const defaultUrl = 'http://www.gailak.com.eg';
const url = process.env.URL || defaultUrl;
app.use(express.static('__dirname'));
app.use(express.json())

app.get('/',(req,res)=>{
    res.send(__dirname + 'index.html')
})

app.post('/',(req,res)=>{
   console.log(req.body)
   const html=` 
   <div class="container">
            <div class="row">
                <div class="col-md-12 d-flex justify-content-center">
                    <div class="border">
                    <p> fullname: <span> ${req.body.namee}</span> </p>
                    <p> email: <span> ${req.body.email}</span> </p>
                    <p> phone: <span> ${req.body.phone}</span> </p>
                    <p> email: <span> ${req.body.subject}</span> </p>
                    <p> message: <span> ${req.body.message}</span> </p>
                    </div>
                </div>
            </div>
        </div>`

    const transporter= nodemailer.createTransport({
        host:'smtp.gmail.com',
        port:465,
        secure:true,
        auth:{
            user:'noreplay@gailak.com.eg',
            pass:'ouabnixgrjqtyles'
        }
    });
     const info=  transporter.sendMail({
        from: req.body.namee,
    to:'noreplay@gailak.com.eg',
    subject:req.body.subject,
    message: req.body.message,
    html:html,
    
     })
    console.log('message send:'+ info.messageId)
  

/*const transporter=   nodemailer.createTransport({
    host:'smtp.gmail.com',
    port:465,
    secure:true,
    auth:{
        user:'Seham.shams20@gmail.com',
        pass:'nenybvmwowvgpgzd'
    }
});
 const infoemail= {
    from:'se <Seham.shams20@gmail.com>' ,
to:req.body.email,
subject:'testing email',
html:html,

 }
 transporter.sendMail(infoemail,(error,info)=>{
if (error){
    console.log(error)
}
else{
    console.log('message send:'+ info.messageId)
    alert('send')
}
 })*/


})
app.use(bodyParser.urlencoded({ extended: true }));
const subscribers = [];
app.post('/export',(req,res)=>{

    const email = req.body.emial;

    if (email) {
        subscribers.push(email);
        console.log(`New subscriber: ${email}`);
       // res.send('Thank you for subscribing to our newsletter!');
    } else {
        res.status(400).send('Email address is required.');
    }
});
app.get('/export-subscribers', (req, res) => {
    // Create a new workbook and add a worksheet
      const filePath = path.join(__dirname, 'subscribers.xlsx');
   
   const workbook = new excel.Workbook();
   const worksheet = workbook.addWorksheet('Subscribers');

    // Add headers to the worksheet
    worksheet.addRow(['Email']);

    // Add subscribers' email addresses to the worksheet
    subscribers.forEach((email) => {
        worksheet.addRow([email]);
    });

    // Generate the Excel file
  
    workbook.xlsx.writeFile(filePath)
        .then(() => {
            console.log('Subscribers list exported to Excel file successfully!');
            res.sendFile(filePath); // Return the Excel file to the user for download
        })
        .catch((err) => {
            console.error('Error generating Excel file:', err);
            res.status(500).send('An error occurred while exporting subscribers data to Excel.');
        });
});
app.listen (port ,()=>{
    console.log(`sever is runing on port ${url}:${port}`)
})