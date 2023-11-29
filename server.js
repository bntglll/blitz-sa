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

app.post("/update", (req, res) => {
  Data.find({})
    .then((data) => {
      data.forEach((item) => {
        Address.findOne({ co_add: item.co_add })
          .then((address) => {
            if (address) {
              Data.updateOne(
                { _id: item._id },
                { co_lat: address.co_lat, co_lng: address.co_lng }
              )
                .then(() => {
                  const googleMapsApiKey =
                    "AIzaSyD4sgjH4RAaAokyujwQO_jSeZDowQ1U9Oo";
                  const googleMapsApiUrl = `https://maps.googleapis.com/maps/api/geocode/json?latlng=${address.co_lat},${address.co_lng}&key=${googleMapsApiKey}`;
                  axios
                    .get(googleMapsApiUrl)
                    .then((response) => {
                      const results = response.data.results;
                      if (results.length > 0) {
                        const location = results[0];
                        Data.updateOne(
                          { _id: item._id },
                          { co_add: location.formatted_address }
                        )
                          .then(() =>
                            console.log("Updated address for " + item.sell_id)
                          )
                          .catch((err) => console.error(err));
                      }
                    })
                    .catch((err) => console.error(err));
                })
                .catch((err) => console.error(err));
            }
          })
          .catch((err) => console.error(err));

        const destination = `${item.coe_prov}, ${item.coe_city}, ${item.coe_count}, ${item.coe_add}`;
        const googleMapsApiKey = "AIzaSyD4sgjH4RAaAokyujwQO_jSeZDowQ1U9Oo";
        const googleMapsApiUrl = `https://maps.googleapis.com/maps/api/geocode/json?address=${destination}&key=${googleMapsApiKey}`;
        axios
          .get(googleMapsApiUrl)
          .then((response) => {
            const results = response.data.results;
            if (results.length > 0) {
              const location = results[0];
              Data.updateOne(
                { _id: item._id },
                {
                  coe_lat: location.geometry.location.lat,
                  coe_lng: location.geometry.location.lng,
                }
              )
                .then(() => {
                  const googleMapsApiUrl = `https://maps.googleapis.com/maps/api/geocode/json?latlng=${location.geometry.location.lat},${location.geometry.location.lng}&key=${googleMapsApiKey}`;
                  axios
                    .get(googleMapsApiUrl)
                    .then((response) => {
                      const results = response.data.results;
                      if (results.length > 0) {
                        const location = results[0];
                        Data.updateOne(
                          { _id: item._id },
                          { coe_add: location.formatted_address }
                        )
                          .then(() =>
                            console.log(
                              "Updated address destination for " + item.sell_id
                            )
                          )
                          .catch((err) => console.error(err));
                      }
                    })
                    .catch((err) => console.error(err));
                })
                .catch((err) => console.error(err));
            }
          })
          .catch((err) => console.error(err));
      });
      res.send("Lokasi berhasil diperbaharui dari MongoDB");
    })
    .catch((err) => console.error(err));
});

app.post("/distance", (req, res) => {
  Data.find({})
    .then((data) => {
      data.forEach((item) => {
        const googleMapsApiKey = "AIzaSyD4sgjH4RAaAokyujwQO_jSeZDowQ1U9Oo";
        const mode = "driving";
        const googleMapsApiUrl = `https://maps.googleapis.com/maps/api/distancematrix/json?origins=${item.co_lat},${item.co_lng}&destinations=${item.coe_lat},${item.coe_lng}&mode=${mode}&key=${googleMapsApiKey}`;

        axios
          .get(googleMapsApiUrl)
          .then((response) => {
            const results = response.data.rows[0].elements[0];

            if (results.status === "OK") {
              const distance = results.distance.text;
              const travelTime = results.duration.text;

              console.log("Distance for " + item.sell_id + ": " + distance);
              console.log(
                "Travel time for " + item.sell_id + ": " + travelTime
              );

              Data.updateOne(
                { _id: item._id },
                { distance: distance, travelTime: travelTime }
              )
                .then(() => {
                  const directionsLink = `https://www.google.com/maps/dir/${item.co_lat},${item.co_lng}/${item.coe_lat},${item.coe_lng}/@${item.co_lat},${item.co_lng},13z/data=!3m1!4b1`;
                  Data.updateOne(
                    { _id: item._id },
                    { directionsLink: directionsLink }
                  )
                    .then(() =>
                      console.log("Updated directions link for " + item.sell_id)
                    )
                    .catch((err) => console.error(err));

                  console.log(
                    "Updated distance, travel time, and directions link for " +
                      item.sell_id
                  );
                })
                .catch((err) => console.error(err));
            }
          })
          .catch((err) => console.error(err));
      });

      res.send(
        "Jarak kilometer, waktu tempuh, dan tautan arah berhasil ditemukan dan disimpan di MongoDB"
      );
    })
    .catch((err) => console.error(err));
});

app.get("/export", (req, res) => {
  Data.find({})
    .then((data) => {
      const exportData = data.map((item) => ({
        Sell_ID: item.sell_id,
        Co_Prov: item.co_prov,
        Co_City: item.co_city,
        Co_Count: item.co_count,
        Co_Add: item.co_add,
        Co_Lat: item.co_lat,
        Co_Lng: item.co_lng,
        Coe_Prov: item.coe_prov,
        Coe_City: item.coe_city,
        Coe_Count: item.coe_count,
        Coe_Add: item.coe_add,
        Coe_Lat: item.coe_lat,
        Coe_Lng: item.coe_lng,
        Assignee: item.assignee,
        Distance: item.distance,
        Travel_Time: item.travelTime,
        Directions_Link: item.directionsLink,
      }));

      const ws = xlsx.utils.json_to_sheet(exportData);
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, "Data");

      const excelFileName = "data.xlsx";
      xlsx.writeFile(wb, excelFileName);

      res.download(excelFileName, (err) => {
        if (!err) {
          fs.unlinkSync(excelFileName);
        }
      });
    })
    .catch((err) => console.error(err));
});

app.listen(3000, () => {
  console.log("Server is running on port 3000");
});
