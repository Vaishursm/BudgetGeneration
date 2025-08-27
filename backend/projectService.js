const db = require('./db');

function createProject({ name, password, start_date, end_date }) {
  const stmt = db.getDB().prepare(`
    INSERT INTO projects (name, password, start_date, end_date)
    VALUES (?, ?, ?, ?)
  `);
  stmt.run(name, password, start_date, end_date);
  return { success: true };
}

function getProjects() {
  return db.getDB().prepare(`SELECT id, name FROM projects`).all();
}

module.exports = { createProject, getProjects };
