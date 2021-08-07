const express = require("express");

const multer = require("multer");

const app = express();

const fileStorageEngine = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, './destination');
    },
    filename: (req, file ,cb) => {
        cb(null, 'uploaded' + '--' + file.originalname);
    },
});

const upload = multer({storage : fileStorageEngine});

app.post('/single',upload.single("excel"),(req,res) => {
    console.log(req.file);
    res.send('single file upload succssful');
});

app.listen(5000);