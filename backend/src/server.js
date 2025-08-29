const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");

const projectRoutes = require("./routes/projects");

const app = express();
app.use(cors());
app.use(bodyParser.json());

// ✅ Mount routes
app.use("/projects", projectRoutes);

const PORT = 5000;
app.listen(PORT, () => {
  console.log(`🚀 Server running on http://localhost:${PORT}`);
});
