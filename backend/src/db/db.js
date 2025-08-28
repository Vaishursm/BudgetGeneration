// db.js
const sqlite3 = require("sqlite3").verbose();
const path = require("path");
console.log("👉 Using database at:", path.resolve("./projects.db"));

const db = new sqlite3.Database("./projects.db", (err) => {
  if (err) {
    console.error("❌ Error opening database", err.message);
  } else {
    console.log("✅ Connected to SQLite database.");

    db.run(
      `CREATE TABLE IF NOT EXISTS projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        projectCode TEXT UNIQUE,
        description TEXT NOT NULL,
        clientName TEXT NOT NULL,
        projectLocation TEXT NOT NULL,
        projectValue REAL NOT NULL,
        startDate TEXT NOT NULL,
        endDate TEXT NOT NULL,
        concreteQty INTEGER NOT NULL,
        fuelCost REAL NOT NULL,
        powerCost REAL NOT NULL,
        filePath TEXT NOT NULL
      )`,
      (err) => {
        if (err) {
          console.error("❌ Error creating table:", err.message);
        } else {
          console.log("✅ Projects table ready.");
        }
      }
    );
  }
});

module.exports = db;
