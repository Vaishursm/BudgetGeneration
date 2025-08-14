const Database = require("better-sqlite3")
const fs = require("fs")
const path = require("path")

const dbPath = path.join(process.cwd(), "budget-generation.sqlite")
const db = new Database(dbPath)

console.log("Setting up database for Budget Generation Software...")

// Enable WAL mode and foreign keys
db.pragma("journal_mode = WAL")
db.pragma("foreign_keys = ON")

// Read and execute the SQL script
const sqlScript = fs.readFileSync(path.join(process.cwd(), "scripts", "03-create-projects-table.sql"), "utf8")
db.exec(sqlScript)

console.log("Database setup completed successfully!")
console.log(`Database created at: ${dbPath}`)

// Close the database connection
db.close()
