const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const mongodb = require('mongodb');
const {google} = require('googleapis');
const fs = require('fs');


const application = express();


//using the middleware of body parser
application.use(bodyParser.urlencoded({extended:true}))

var storage = multer.diskStorage({
    destination:function(req,file, cb){
        cb(null,'uploads')
    },
    filename:function(req,file,cb){
        cb(null,file.fieldname + path.extname(file.originalname))
    }
})

var upload = multer({storage:storage})


//configuring mongodb

const mongoClient = mongodb.MongoClient;
const url = 'mongodb://localhost:27017';

mongoClient.connect(url,{
    useUnifiedTopology:true,
    useNewUrlParser:true
},(err,client)=>{
    if(err){
        return console.log(err);
    }
    db = client.db('Excel');
    application.listen(3000,()=>{
        console.log("listenning at 3k");
    })
})


// configuring the upload file route

application.post('/uploadFile',upload.single('excelFile'),(req,res,next)=> {
    const file = req.file;
    if(!file){
        const error = new Error("retry");
        error.httpStatusCode = 400;
        return next(error);
    }
    res.send(file);
})


//configuring the home route

application.get('/',(req,res)=>{
    res.sendFile(__dirname + '/main.html');
})



application.listen(5000,()=>{
    console.log("server is listenning on port 5k");
})



// using google drive api 

const client_ID = '350646417561kogp4feor1raeo0k3hkn.apps.googleusercontent.com';
const client_SECREAT = 'WxcYB0AdIS5rIxYjrHwM';
const Redirect_URI = 'https://developers.google.com/oauthplayground';

const ref_token = '1//04vNKLJZOozqeCgYIARAAGAQSNwFrF48Fa991crsGKgmJ2laMFJXJZ1Ni6cjaMXLTMSIWjLGmJb_KGzw';

const oauth2Client = new google.auth.OAuth2(
    client_ID,
    client_SECREAT,
    Redirect_URI
);

oauth2Client.setCredentials({refresh_token:ref_token});

const gDrive = google.drive({
    version:'v3',
    auth: oauth2Client
})

const file_Path = path.join(__dirname + '/uploads/excelFile.xlsx');

async function uploadFile(){
    try{
        const response = await gDrive.files.create({
            requestBody:{
                name:'excelFile.xlsx',
                mineType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            },
            media: {
                mineType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                body: fs.createReadStream(file_Path)
            }
        })

        console.log(response.data);

    }catch (error){
        console.log(error.message);
    }
}


setTimeout(() => { uploadFile(); }, 60000);
