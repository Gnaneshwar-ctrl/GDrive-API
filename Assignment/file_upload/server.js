const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const path = require('path');
const excelToJson = require('convert-excel-to-json');
const mongodb = require('mongodb');
const mongoose = require('mongoose');
const {google} = require('googleapis');
const fs = require('fs');


const application = express();


//using the middleware of body parser
application.use(bodyParser.urlencoded({extended:true}))

const storage = multer.diskStorage({
    destination:function(req,file, cb){
        cb(null,'uploads')
    },
    filename:function(req,file,cb){
        cb(null,file.fieldname + path.extname(file.originalname))
    }
})

const upload = multer({storage:storage})

//configuring mongodb

const mongoClient = mongodb.MongoClient;
const url = 'mongodb://localhost:27017/Excel';

mongoose.connect(url,{
    useUnifiedTopology:true,
    useNewUrlParser:true
});

const db = mongoose.connection;

db.once('open',function(){
    console.log("db connected")
})

// -> Read Excel File to Json Data

const excelData = excelToJson({
    sourceFile: '../myexcel.xlsx',
    sheets:[{
		// Excel Sheet Name
        name: 'CUS OUTBOUND RAW DATA',
		
		// Header Row -> be skipped and will not be present at our result object.
		header:{
            rows: 1
        },
		
		// Mapping columns to keys
        columnToKey: {
        	A: 'DATE',
    		B: 'MODE', 
			C: 'LOCATION',
			D: 'CUSTOMER',
            E: 'PRODUCT CODE',
            F: 'SOURCE',
            G: 'RAILCAR',
            H: 'FLEET',
            I: 'SUBFLEET',
            J: 'RAILCAR SEALS',
            K: 'BOL',
            L: 'TERMINAL / DESTINATION',
            M: 'CITY',
            N: 'STATE',
            O: 'WEIGHT',
            P: 'TEMPERATURE',
            Q: 'DENSITY',
            R: 'S&W %',
            S: 'S&W',
            T: 'NET OIL',
            U: 'TOTAL VOL',
            V: 'S&W',
            W: 'NET OIL',
            X: 'TOTAL VOL',
            Y: 'BOL DATE',
            Z: 'HEEL VOLUME',
            AA: 'HEEL WEIGHT',
            AB: 'Contract Id'
        }
    }]
});

// -> Log Excel Data to Console
console.log(excelData);

const inBoundSchema =  new mongoose.Schema({
    date:{
        type:Date,
        required:true,
    },
    mode:{
        type:String,
        required:true,
        trim:true
    },
    location:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    customer:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    product_code: {
        type:String,
        required:true,
        trim:true
    },
    source:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    railcar:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    fleet:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    subfleet:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    railcar_seals:{
        type:[String]
    },
    bol:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    destination:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    city:{
        type: String,
        required: true,
        trim:true,
    },
    state:{
        type: String,
        required: true,
        trim:true,
        uppercase:true
    },
    weight:{
        type:Number,
        required:true
    },
    temperature:{
        type:Number,
        required:true
    },
    density:{
        type:Number,
        required:true
    },
    sw_percent:{
        type:Number,
        required:true
    },
    sw_bbl:{
        type:Number,
        required:true
    },
    net_oil_bbl:{
        type:Number,
        required:true
    },
    total_volume_bbl:{
        type:Number,
        required:true
    },
    sw_m3:{
        type:Number,
        required:true
    },
    net_oil_m3:{
        type:Number,
        required:true
    },
    total_volume_m3:{
        type:Number,
        required:true
    },
    bol_date:{
        type:Date,
        required:true
    },
    heel_volume:{
        type:Number,
        required:true
    },
    heel_weight:{
        type:Number,
        required:true
    }
})

const InBound = mongoose.model('InBound',inBoundSchema)

// ADD 

const dd = require("date-and-time");

const toDBOutObject = (record) => {
  const res = {
    date: strToDate(record["A"]),
    mode: record["B"],
    location: record["C"],
    customer: record["D"],
    product_code: record["E"],
    source: record["F"],
    railcar: record["G"],
    fleet: record["H"],
    subfleet: record["I"],
    railcar_seals: strToArr(record["J"]),
    bol: record["K"],
    destination: record["L"],
    city: record["M"],
    state: record["N"],
    weight: strToNum(record["O"] + ""),
    temperature: strToNum(record["P"] + ""),
    density: strToNum(record["Q"] + ""),
    sw_percent: strToNum(record["R"] + ""),
    sw_bbl: strToNum(record["S"] + ""),
    net_oil_bbl: strToNum(record["T"] + ""),
    total_volume_bbl: strToNum(record["U"] + ""),
    sw_m3: strToNum(record["V"] + ""),
    net_oil_m3: strToNum(record["W"] + ""),
    total_volume_m3: strToNum(record["X"] + ""),
    bol_date: strToDate(record["Y"]),
    heel_volume: strToNum(record["Z"] + ""),
    heel_weight: strToNum(record["AA"] + ""),
    contract_id: strToNum(record["AB"] + ""),
  };
  return res;
};
const toDBInObject = (record) => {
  const res = {
    date: strToDate(record["A"]),
    mode: record["B"],
    location: record["C"],
    customer: record["D"],
    product_code: record["E"],
    source: record["F"],
    railcar: record["G"],
    fleet: record["H"],
    subfleet: record["I"],
    railcar_seals: strToArr(record["J"]),
    bol: record["K"],
    destination: record["L"],
    city: record["M"],
    state: record["N"],
    weight: strToNum(record["O"] + ""),
    temperature: strToNum(record["P"] + ""),
    density: strToNum(record["Q"] + ""),
    sw_percent: strToNum(record["R"] + ""),
    sw_bbl: strToNum(record["S"] + ""),
    net_oil_bbl: strToNum(record["T"] + ""),
    total_volume_bbl: strToNum(record["U"] + ""),
    sw_m3: strToNum(record["V"] + ""),
    net_oil_m3: strToNum(record["W"] + ""),
    total_volume_m3: strToNum(record["X"] + ""),
    bol_date: strToDate(record["Y"]),
    heel_volume: strToNum(record["Z"] + ""),
    heel_weight: strToNum(record["AA"] + ""),
  };
  return res;
};

const strToDate = (str) => {
  date1 = new Date(str);
  return dd.format(date1, "MMM DD YYYY");
};

const strToArr = (str) => {
  return str.split();
};

const strToNum = (str) => {
  return str.replace(/[^\d.-]/g, "");
};


// END 

mongoClient.connect(url, function(err, client) {
    if (err) throw err;
    var dbo = client.db("Excel");
        dbo.collection("customers").insertOne(excelData, function(err, res) {
          if (err) throw err;
          console.log("1 document inserted");
          res.send;
          db.close();
    });
  });

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

const client_ID = '350646417561-o8uitid83rflkogp4feor1raeo0k3hkn.apps.googleusercontent.com';
const client_SECREAT = 'WxcYB0AdIS5rIxYjrHwbQoyM';
const Redirect_URI = 'https://developers.google.com/oauthplayground';

const ref_token = '1//04vNKLJZOozqeCgYIARAAGAQSNwF-L9IrS1mNnyV1nm1MPIkrF48Fa991crsGKgmJ2laMFJXJZ1Ni6cjaMXLTMSIWjLGmJb_KGzw';

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


//setTimeout(() => { uploadFile(); }, 60000);

