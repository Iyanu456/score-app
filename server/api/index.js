require("dotenv").config();
const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const upload_route = require("./routes/upload");
const path = require('path');


const app = express();
app.use(bodyParser.json());


app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "PUT", "DELETE", "PATCH", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

// Serve static files from the 'dist' folder
app.use(express.static(path.join(__dirname, '..', '..', 'client', 'dist')));

// Any routes you have for API endpoints
// app.use('/api', apiRouter);

// Catch-all route to serve the Vite frontend
app.get('*', (req, res) => {
  res.sendFile(path.resolve(__dirname, '..', '..', 'client', 'dist', 'index.html'));
});


app.use("/api", upload_route);




const PORT = process.env.PORT || 3000;
app.listen(PORT, console.log(`Server listening on port ${PORT}`));