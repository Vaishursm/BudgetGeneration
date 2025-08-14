const { app } = require('electron');
const Database = require('better-sqlite3');
const path = require('path');
let db;

function init() {
  const dbPath = path.join(app.getPath('userData'), 'budget.db');
  db = new Database(dbPath);
  db.prepare(`
    CREATE TABLE IF NOT EXISTS projects (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT,
      password TEXT,
      start_date TEXT,
      end_date TEXT
    )
  `).run();
}

function getDB() {
  return db;
}

module.exports = { init, getDB };
