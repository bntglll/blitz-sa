const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const mongoose = require("mongoose");
const axios = require("axios");
const fs = require("fs");

const app = express();

app.set("view engine", "ejs");
app.use(express.static("public"));

// Menggunakan memory storage untuk multer agar file tidak disimpan di disk
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

const username = "myangkasasa";
const password = "GxAJRKhHelscM5gM";
const url = `mongodb+srv://${username}:${password}@cluster0.nrt1pdw.mongodb.net/`;
const dbName = "blitz";
const collectionName = "data";

mongoose
  .connect(url + dbName, { useNewUrlParser: true, useUnifiedTopology: true })
  .then(() => console.log("Connected to MongoDB"))
  .catch((err) => console.error(err));

const dataSchema = new mongoose.Schema({
  sell_id: String,
  co_prov: String,
  co_city: String,
  co_count: String,
  co_add: String,
  coe_prov: String,
  coe_city: String,
  coe_count: String,
  coe_add: String,
  assignee: String,
  co_lat: Number,
  co_lng: Number,
  coe_lat: Number,
  coe_lng: Number,
  distance: String,
  directionsLink: String,
  travelTime: String,
});

const Data = mongoose.model(collectionName, dataSchema);

const addressSchema = new mongoose.Schema({
  co_add: String,
  co_lat: Number,
  co_lng: Number,
});

const Address = mongoose.model("addresses", addressSchema);

app.get("/", (req, res) => {
  Data.find({})
    .then((data) => {
      res.render("index", { data: data });
    })
    .catch((err) => console.error(err));
});

app.post("/upload", upload.single("excel"), (req, res) => {
  // Membaca file dari buffer yang diberikan oleh multer
  const workbook = xlsx.read(req.file.buffer);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet);

  Data.deleteMany({})
    .then(() => console.log("Deleted old data"))
    .catch((err) => console.error(err));

  Data.insertMany(data)
    .then(() => console.log("Inserted new data"))
    .catch((err) => console.error(err));

  res.send("File excel berhasil diunggah dan disimpan ke MongoDB");
});

app.listen(3001, () => {
  console.log("Server is running on port 3001");
});
