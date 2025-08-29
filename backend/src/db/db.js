// src/db/db.js
const { Sequelize, DataTypes } = require("sequelize");

// Create Sequelize instance (SQLite file will be projects.db)
const sequelize = new Sequelize({
  dialect: "sqlite",
  storage: "./projects.db",
  logging: false, // disable SQL logs
  timestamps: false,
});

// Define Project model
const Project = sequelize.define("Project", {
  projectCode: { type: DataTypes.STRING, unique: true, allowNull: false },
  description: { type: DataTypes.STRING, allowNull: false },
  clientName: { type: DataTypes.STRING, allowNull: false },
  projectLocation: { type: DataTypes.STRING, allowNull: false },
  projectValue: { type: DataTypes.FLOAT, allowNull: false },
  startDate: { type: DataTypes.STRING, allowNull: false },
  endDate: { type: DataTypes.STRING, allowNull: false },
  concreteQty: { type: DataTypes.INTEGER, allowNull: false },
  fuelCost: { type: DataTypes.FLOAT, allowNull: false },
  powerCost: { type: DataTypes.FLOAT, allowNull: false },
  filePath: { type: DataTypes.STRING, allowNull: false },
  password: { type: DataTypes.STRING, allowNull: false },
});

// Sync database (creates tables if not exists)
sequelize.sync();

module.exports = { sequelize, Project };
