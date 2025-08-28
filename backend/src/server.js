  // server.js
  const express = require("express");
  const bodyParser = require("body-parser");
  const cors = require("cors");
  const db = require("./db/db");

  const app = express();
  app.use(cors());
  app.use(bodyParser.json());

  // ✅ Create new project
  app.post("/projects", (req, res) => {
    const {
      projectCode,
      description,
      clientName,
      projectLocation,
      projectValue,
      startDate,
      endDate,
      concreteQty,
      fuelCost,
      powerCost,
      filePath,
    } = req.body;

    const sql = `INSERT INTO projects 
      (projectCode, description, clientName, projectLocation, projectValue, startDate, endDate, concreteQty, fuelCost, powerCost, filePath)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

    db.run(
      sql,
      [
        projectCode,
        description,
        clientName,
        projectLocation,
        projectValue,
        startDate,
        endDate,
        concreteQty,
        fuelCost,
        powerCost,
        filePath,
      ],
      function (err) {
        if (err) {
          if (err.code === "SQLITE_CONSTRAINT") {
            res.status(400).json({ error: "Project already exists" });
          } else {
            res.status(500).json({ error: err.message });
          }
        } else {
          res.json({ id: this.lastID, message: "Project created successfully" });
        }
      }
    );
  });

  // ✅ Get all projects
  app.get("/projects", (req, res) => {
    db.all("SELECT * FROM projects", [], (err, rows) => {
      if (err) {
        res.status(500).json({ error: err.message });
      } else {
        res.json(rows);
      }
    });
  });

  // ✅ Get project by ID
  app.get("/projects/:id", (req, res) => {
    db.get("SELECT * FROM projects WHERE id = ?", [req.params.id], (err, row) => {
      if (err) {
        res.status(500).json({ error: err.message });
      } else if (!row) {
        res.status(404).json({ error: "Project not found" });
      } else {
        res.json(row);
      }
    });
  });

  // ✅ Update project
  // ✅ Update project (with unique projectCode check)
  app.put("/projects/:id", (req, res) => {
    const {
      projectCode,
      description,
      clientName,
      projectLocation,
      projectValue,
      startDate,
      endDate,
      concreteQty,
      fuelCost,
      powerCost,
      filePath,
    } = req.body;

    // Check if projectCode already exists in another record
    db.get(
      "SELECT id FROM projects WHERE projectCode = ? AND id != ?",
      [projectCode, req.params.id],
      (err, row) => {
        if (err) return res.status(500).json({ error: err.message });

        if (row) {
          return res.status(400).json({ error: "Project code already exists" });
        }

        const sql = `UPDATE projects SET 
          projectCode = ?, description = ?, clientName = ?, projectLocation = ?, projectValue = ?, 
          startDate = ?, endDate = ?, concreteQty = ?, fuelCost = ?, powerCost = ?, filePath = ?
          WHERE id = ?`;

        db.run(
          sql,
          [
            projectCode,
            description,
            clientName,
            projectLocation,
            projectValue,
            startDate,
            endDate,
            concreteQty,
            fuelCost,
            powerCost,
            filePath,
            req.params.id,
          ],
          function (err) {
            if (err) {
              res.status(500).json({ error: err.message });
            } else if (this.changes === 0) {
              res.status(404).json({ error: "Project not found" });
            } else {
              res.json({ message: "Project updated successfully" });
            }
          }
        );
      }
    );
  });


  // ✅ Delete project
  app.delete("/projects/:id", (req, res) => {
    db.run("DELETE FROM projects WHERE id = ?", [req.params.id], function (err) {
      if (err) {
        res.status(500).json({ error: err.message });
      } else if (this.changes === 0) {
        res.status(404).json({ error: "Project not found" });
      } else {
        res.json({ message: "Project deleted successfully" });
      }
    });
  });

  const PORT = 5000;
  app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
  });
