import Database from "better-sqlite3"
import { join } from "path"

// Singleton database connection
let db: Database.Database | null = null

export function getDatabase() {
  if (!db) {
    const dbPath = join(process.cwd(), "budget-generation.sqlite")
    db = new Database(dbPath)
    db.pragma("foreign_keys = ON")

    createProjectsTableIfNotExists()
  }
  return db
}

function createProjectsTableIfNotExists() {
  if (!db) return

  const tableExists = db
    .prepare(`
    SELECT name FROM sqlite_master 
    WHERE type='table' AND name='projects'
  `)
    .get()

  if (!tableExists) {
    db.exec(`
      CREATE TABLE projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_name TEXT NOT NULL UNIQUE,
        project_description TEXT,
        client_name TEXT NOT NULL,
        project_location TEXT NOT NULL,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        workbook_name TEXT NOT NULL,
        workbook_location TEXT NOT NULL,
        password_hash TEXT NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
      );
      
      CREATE INDEX idx_projects_name ON projects(project_name);
      CREATE INDEX idx_projects_created_at ON projects(created_at);
      
      CREATE TRIGGER update_projects_updated_at 
      AFTER UPDATE ON projects
      BEGIN
        UPDATE projects SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;
    `)
  }

  createEquipmentTablesIfNotExists()
}

function createEquipmentTablesIfNotExists() {
  if (!db) return

  // Check if equipment tables exist
  const equipmentTableExists = db
    .prepare(`SELECT name FROM sqlite_master WHERE type='table' AND name='equipment_master'`)
    .get()

  if (!equipmentTableExists) {
    db.exec(`
      -- Master equipment table with all available equipment
      CREATE TABLE equipment_master (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        category TEXT NOT NULL,
        equipment_name TEXT NOT NULL,
        unit TEXT NOT NULL,
        rate REAL NOT NULL,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
      );

      -- Project equipment selections and data
      CREATE TABLE project_equipment (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        equipment_id INTEGER NOT NULL,
        is_selected BOOLEAN DEFAULT FALSE,
        mob_date TEXT,
        demob_date TEXT,
        quantity REAL DEFAULT 1,
        hours_month REAL,
        depreciation_percent REAL,
        shifts INTEGER,
        hire_charges REAL,
        purchase_value REAL,
        cost REAL,
        remarks TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
        FOREIGN KEY (equipment_id) REFERENCES equipment_master(id)
      );

      -- Electricals tab data
      CREATE TABLE project_electricals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        installation_percent REAL DEFAULT 0,
        hv_cables REAL DEFAULT 0,
        lv_cables REAL DEFAULT 0,
        control_cables REAL DEFAULT 0,
        earthing REAL DEFAULT 0,
        lighting REAL DEFAULT 0,
        others REAL DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
      );

      -- Pipeline expenses data
      CREATE TABLE project_pipeline_expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        quantity REAL DEFAULT 0,
        cost_per_unit REAL DEFAULT 0,
        amount REAL DEFAULT 0,
        remarks TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
      );

      -- Elect/Mechanic cost data
      CREATE TABLE project_mechanic_costs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        nos INTEGER DEFAULT 0,
        salary_per_month REAL DEFAULT 0,
        no_of_months INTEGER DEFAULT 0,
        salary_cost REAL DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
      );

      -- Misc & Non-ERP expenses data
      CREATE TABLE project_misc_expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        amount REAL DEFAULT 0,
        remarks TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
      );

      -- Staff salary data
      CREATE TABLE project_staff_salary (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        category TEXT NOT NULL,
        nos INTEGER DEFAULT 0,
        salary_per_month REAL DEFAULT 0,
        no_of_months INTEGER DEFAULT 0,
        salary_cost REAL DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
      );

      -- Create indexes for better performance
      CREATE INDEX idx_project_equipment_project_id ON project_equipment(project_id);
      CREATE INDEX idx_project_equipment_equipment_id ON project_equipment(equipment_id);
      CREATE INDEX idx_project_electricals_project_id ON project_electricals(project_id);
      CREATE INDEX idx_project_pipeline_project_id ON project_pipeline_expenses(project_id);
      CREATE INDEX idx_project_mechanic_project_id ON project_mechanic_costs(project_id);
      CREATE INDEX idx_project_misc_project_id ON project_misc_expenses(project_id);
      CREATE INDEX idx_project_staff_project_id ON project_staff_salary(project_id);

      -- Create triggers for updated_at timestamps
      CREATE TRIGGER update_project_equipment_updated_at 
      AFTER UPDATE ON project_equipment
      BEGIN
        UPDATE project_equipment SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;

      CREATE TRIGGER update_project_electricals_updated_at 
      AFTER UPDATE ON project_electricals
      BEGIN
        UPDATE project_electricals SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;

      CREATE TRIGGER update_project_pipeline_updated_at 
      AFTER UPDATE ON project_pipeline_expenses
      BEGIN
        UPDATE project_pipeline_expenses SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;

      CREATE TRIGGER update_project_mechanic_updated_at 
      AFTER UPDATE ON project_mechanic_costs
      BEGIN
        UPDATE project_mechanic_costs SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;

      CREATE TRIGGER update_project_misc_updated_at 
      AFTER UPDATE ON project_misc_expenses
      BEGIN
        UPDATE project_misc_expenses SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;

      CREATE TRIGGER update_project_staff_updated_at 
      AFTER UPDATE ON project_staff_salary
      BEGIN
        UPDATE project_staff_salary SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
      END;
    `)

    insertSampleEquipmentData()
  }
}

function insertSampleEquipmentData() {
  if (!db) return

  const equipmentData = [
    // Major Concrete
    { category: "Major Concrete", equipment_name: "Concrete Mixer 10/7", unit: "No", rate: 25000 },
    { category: "Major Concrete", equipment_name: "Concrete Pump", unit: "No", rate: 35000 },
    { category: "Major Concrete", equipment_name: "Transit Mixer", unit: "No", rate: 28000 },
    { category: "Major Concrete", equipment_name: "Batching Plant", unit: "No", rate: 45000 },

    // Major Conveyance
    { category: "Major Conveyance", equipment_name: "Belt Conveyor", unit: "Mtr", rate: 1500 },
    { category: "Major Conveyance", equipment_name: "Bucket Elevator", unit: "No", rate: 18000 },
    { category: "Major Conveyance", equipment_name: "Screw Conveyor", unit: "Mtr", rate: 2000 },

    // Major Crane
    { category: "Major Crane", equipment_name: "Mobile Crane 25T", unit: "No", rate: 32000 },
    { category: "Major Crane", equipment_name: "Tower Crane", unit: "No", rate: 55000 },
    { category: "Major Crane", equipment_name: "Crawler Crane 50T", unit: "No", rate: 48000 },

    // Major DG Sets
    { category: "Major DG Sets", equipment_name: "DG Set 125 KVA", unit: "No", rate: 22000 },
    { category: "Major DG Sets", equipment_name: "DG Set 250 KVA", unit: "No", rate: 38000 },
    { category: "Major DG Sets", equipment_name: "DG Set 500 KVA", unit: "No", rate: 65000 },

    // Major Material Handling
    { category: "Major Material Handling", equipment_name: "Forklift 3T", unit: "No", rate: 15000 },
    { category: "Major Material Handling", equipment_name: "Wheel Loader", unit: "No", rate: 28000 },
    { category: "Major Material Handling", equipment_name: "Excavator", unit: "No", rate: 35000 },

    // Major Non-Concrete
    { category: "Major Non-Concrete", equipment_name: "Compressor 750 CFM", unit: "No", rate: 18000 },
    { category: "Major Non-Concrete", equipment_name: "Welding Machine", unit: "No", rate: 8000 },
    { category: "Major Non-Concrete", equipment_name: "Cutting Machine", unit: "No", rate: 12000 },

    // Major Others
    { category: "Major Others", equipment_name: 'Water Pump 6"', unit: "No", rate: 5000 },
    { category: "Major Others", equipment_name: "Dewatering Pump", unit: "No", rate: 8000 },
    { category: "Major Others", equipment_name: "Submersible Pump", unit: "No", rate: 6000 },

    // Minor E Equipments
    { category: "Minor E Equipments", equipment_name: "Drill Machine", unit: "No", rate: 2500 },
    { category: "Minor E Equipments", equipment_name: 'Grinder 4"', unit: "No", rate: 1800 },
    { category: "Minor E Equipments", equipment_name: "Circular Saw", unit: "No", rate: 3200 },

    // Hired Equipments
    { category: "Hired Equipments", equipment_name: "Hired Crane", unit: "Hour", rate: 1200 },
    { category: "Hired Equipments", equipment_name: "Hired Excavator", unit: "Hour", rate: 800 },
    { category: "Hired Equipments", equipment_name: "Hired Truck", unit: "Hour", rate: 600 },

    // Fixed Exp. - Tower Crane
    { category: "Fixed Exp. - Tower Crane", equipment_name: "Tower Crane Foundation", unit: "LS", rate: 150000 },
    { category: "Fixed Exp. - Tower Crane", equipment_name: "Tower Crane Erection", unit: "LS", rate: 80000 },
    { category: "Fixed Exp. - Tower Crane", equipment_name: "Tower Crane Dismantling", unit: "LS", rate: 60000 },

    // Fixed Exp. - BP Related
    { category: "Fixed Exp. - BP Related", equipment_name: "Batching Plant Foundation", unit: "LS", rate: 120000 },
    { category: "Fixed Exp. - BP Related", equipment_name: "Batching Plant Erection", unit: "LS", rate: 100000 },
    { category: "Fixed Exp. - BP Related", equipment_name: "Aggregate Bins", unit: "LS", rate: 75000 },

    // Lighting/Single Phase Equips
    { category: "Lighting/Single Phase Equips", equipment_name: "LED Flood Light 100W", unit: "No", rate: 2500 },
    { category: "Lighting/Single Phase Equips", equipment_name: "Street Light", unit: "No", rate: 3500 },
    { category: "Lighting/Single Phase Equips", equipment_name: "Emergency Light", unit: "No", rate: 1500 },
  ]

  const stmt = db.prepare(`
    INSERT INTO equipment_master (category, equipment_name, unit, rate)
    VALUES (?, ?, ?, ?)
  `)

  equipmentData.forEach((equipment) => {
    stmt.run(equipment.category, equipment.equipment_name, equipment.unit, equipment.rate)
  })
}

// Database types
export interface User {
  id: number
  email: string
  username: string
  password_hash: string
  first_name?: string
  last_name?: string
  created_at: string
  updated_at: string
}

export interface Post {
  id: number
  title: string
  content: string
  author_id: number
  status: "draft" | "published" | "archived"
  created_at: string
  updated_at: string
  author?: {
    username: string
    first_name?: string
    last_name?: string
  }
  categories?: Category[]
}

export interface Category {
  id: number
  name: string
  description?: string
  created_at: string
}

export interface Project {
  id: number
  project_name: string
  project_description?: string
  client_name: string
  project_location: string
  start_date: string
  end_date: string
  workbook_name: string
  workbook_location: string
  password_hash: string
  created_at: string
  updated_at: string
}

// Equipment interfaces
export interface EquipmentMaster {
  id: number
  category: string
  equipment_name: string
  unit: string
  rate: number
  created_at: string
}

export interface ProjectEquipment {
  id: number
  project_id: number
  equipment_id: number
  is_selected: boolean
  mob_date?: string
  demob_date?: string
  quantity: number
  hours_month?: number
  depreciation_percent?: number
  shifts?: number
  hire_charges?: number
  purchase_value?: number
  cost?: number
  remarks?: string
  created_at: string
  updated_at: string
  equipment?: EquipmentMaster
}

export interface ProjectElectricals {
  id: number
  project_id: number
  installation_percent: number
  hv_cables: number
  lv_cables: number
  control_cables: number
  earthing: number
  lighting: number
  others: number
  created_at: string
  updated_at: string
}

export interface ProjectPipelineExpense {
  id: number
  project_id: number
  category: string
  quantity: number
  cost_per_unit: number
  amount: number
  remarks?: string
  created_at: string
  updated_at: string
}

export interface ProjectMechanicCost {
  id: number
  project_id: number
  category: string
  nos: number
  salary_per_month: number
  no_of_months: number
  salary_cost: number
  created_at: string
  updated_at: string
}

export interface ProjectMiscExpense {
  id: number
  project_id: number
  category: string
  amount: number
  remarks?: string
  created_at: string
  updated_at: string
}

export interface ProjectStaffSalary {
  id: number
  project_id: number
  category: string
  nos: number
  salary_per_month: number
  no_of_months: number
  salary_cost: number
  created_at: string
  updated_at: string
}

// Database operations
export class DatabaseOperations {
  private db: Database.Database

  constructor() {
    this.db = getDatabase()
  }

  // User operations
  getUserByEmail(email: string): User | undefined {
    return this.db.prepare("SELECT * FROM users WHERE email = ?").get(email) as User | undefined
  }

  getUserById(id: number): User | undefined {
    return this.db.prepare("SELECT * FROM users WHERE id = ?").get(id) as User | undefined
  }

  createUser(userData: Omit<User, "id" | "created_at" | "updated_at">): User {
    const stmt = this.db.prepare(`
      INSERT INTO users (email, username, password_hash, first_name, last_name)
      VALUES (?, ?, ?, ?, ?)
    `)
    const result = stmt.run(
      userData.email,
      userData.username,
      userData.password_hash,
      userData.first_name,
      userData.last_name,
    )
    return this.getUserById(result.lastInsertRowid as number)!
  }

  // Post operations
  getAllPosts(status?: string): Post[] {
    let query = `
      SELECT p.*, u.username, u.first_name, u.last_name
      FROM posts p
      JOIN users u ON p.author_id = u.id
    `
    const params: any[] = []

    if (status) {
      query += " WHERE p.status = ?"
      params.push(status)
    }

    query += " ORDER BY p.created_at DESC"

    return this.db.prepare(query).all(...params) as Post[]
  }

  getPostById(id: number): Post | undefined {
    const post = this.db
      .prepare(`
      SELECT p.*, u.username, u.first_name, u.last_name
      FROM posts p
      JOIN users u ON p.author_id = u.id
      WHERE p.id = ?
    `)
      .get(id) as Post | undefined

    if (post) {
      // Get categories for this post
      const categories = this.db
        .prepare(`
        SELECT c.* FROM categories c
        JOIN post_categories pc ON c.id = pc.category_id
        WHERE pc.post_id = ?
      `)
        .all(id) as Category[]
      post.categories = categories
    }

    return post
  }

  createPost(postData: Omit<Post, "id" | "created_at" | "updated_at">): Post {
    const stmt = this.db.prepare(`
      INSERT INTO posts (title, content, author_id, status)
      VALUES (?, ?, ?, ?)
    `)
    const result = stmt.run(postData.title, postData.content, postData.author_id, postData.status)
    return this.getPostById(result.lastInsertRowid as number)!
  }

  updatePost(id: number, postData: Partial<Omit<Post, "id" | "created_at" | "updated_at">>): Post | undefined {
    const updates: string[] = []
    const values: any[] = []

    Object.entries(postData).forEach(([key, value]) => {
      if (value !== undefined && key !== "id" && key !== "created_at" && key !== "updated_at") {
        updates.push(`${key} = ?`)
        values.push(value)
      }
    })

    if (updates.length === 0) return this.getPostById(id)

    updates.push("updated_at = CURRENT_TIMESTAMP")
    values.push(id)

    const stmt = this.db.prepare(`UPDATE posts SET ${updates.join(", ")} WHERE id = ?`)
    stmt.run(...values)

    return this.getPostById(id)
  }

  deletePost(id: number): boolean {
    const stmt = this.db.prepare("DELETE FROM posts WHERE id = ?")
    const result = stmt.run(id)
    return result.changes > 0
  }

  // Category operations
  getAllCategories(): Category[] {
    return this.db.prepare("SELECT * FROM categories ORDER BY name").all() as Category[]
  }

  getCategoryById(id: number): Category | undefined {
    return this.db.prepare("SELECT * FROM categories WHERE id = ?").get(id) as Category | undefined
  }

  // Project operations
  getAllProjects(): Project[] {
    return this.db
      .prepare(
        "SELECT id, project_name, client_name, project_location, start_date, end_date, created_at FROM projects ORDER BY created_at DESC",
      )
      .all() as Project[]
  }

  getProjectById(id: number): Project | undefined {
    return this.db.prepare("SELECT * FROM projects WHERE id = ?").get(id) as Project | undefined
  }

  getProjectByName(projectName: string): Project | undefined {
    return this.db.prepare("SELECT * FROM projects WHERE project_name = ?").get(projectName) as Project | undefined
  }

  createProject(projectData: Omit<Project, "id" | "created_at" | "updated_at">): Project {
    const stmt = this.db.prepare(`
      INSERT INTO projects (project_name, project_description, client_name, project_location, start_date, end_date, workbook_name, workbook_location, password_hash)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `)
    const result = stmt.run(
      projectData.project_name,
      projectData.project_description,
      projectData.client_name,
      projectData.project_location,
      projectData.start_date,
      projectData.end_date,
      projectData.workbook_name,
      projectData.workbook_location,
      projectData.password_hash,
    )
    return this.getProjectById(result.lastInsertRowid as number)!
  }

  updateProject(
    id: number,
    projectData: Partial<Omit<Project, "id" | "created_at" | "updated_at">>,
  ): Project | undefined {
    const updates: string[] = []
    const values: any[] = []

    Object.entries(projectData).forEach(([key, value]) => {
      if (value !== undefined && key !== "id" && key !== "created_at" && key !== "updated_at") {
        updates.push(`${key} = ?`)
        values.push(value)
      }
    })

    if (updates.length === 0) return this.getProjectById(id)

    updates.push("updated_at = CURRENT_TIMESTAMP")
    values.push(id)

    const stmt = this.db.prepare(`UPDATE projects SET ${updates.join(", ")} WHERE id = ?`)
    stmt.run(...values)

    return this.getProjectById(id)
  }

  verifyProjectPassword(projectName: string, passwordHash: string): Project | undefined {
    return this.db
      .prepare("SELECT * FROM projects WHERE project_name = ? AND password_hash = ?")
      .get(projectName, passwordHash) as Project | undefined
  }

  // Equipment Master operations
  getEquipmentByCategory(category: string): EquipmentMaster[] {
    return this.db
      .prepare("SELECT * FROM equipment_master WHERE category = ? ORDER BY equipment_name")
      .all(category) as EquipmentMaster[]
  }

  getAllEquipmentCategories(): string[] {
    const result = this.db.prepare("SELECT DISTINCT category FROM equipment_master ORDER BY category").all() as {
      category: string
    }[]
    return result.map((r) => r.category)
  }

  // Project Equipment operations
  getProjectEquipment(projectId: number, category?: string): ProjectEquipment[] {
    let query = `
      SELECT pe.*, em.category, em.equipment_name, em.unit, em.rate
      FROM project_equipment pe
      JOIN equipment_master em ON pe.equipment_id = em.id
      WHERE pe.project_id = ?
    `
    const params: any[] = [projectId]

    if (category) {
      query += " AND em.category = ?"
      params.push(category)
    }

    query += " ORDER BY em.equipment_name"

    return this.db.prepare(query).all(...params) as ProjectEquipment[]
  }

  upsertProjectEquipment(projectId: number, equipmentId: number, data: Partial<ProjectEquipment>): ProjectEquipment {
    const existing = this.db
      .prepare("SELECT * FROM project_equipment WHERE project_id = ? AND equipment_id = ?")
      .get(projectId, equipmentId) as ProjectEquipment | undefined

    if (existing) {
      // Update existing record
      const updates: string[] = []
      const values: any[] = []

      Object.entries(data).forEach(([key, value]) => {
        if (
          value !== undefined &&
          key !== "id" &&
          key !== "project_id" &&
          key !== "equipment_id" &&
          key !== "created_at" &&
          key !== "updated_at"
        ) {
          updates.push(`${key} = ?`)
          values.push(value)
        }
      })

      if (updates.length > 0) {
        updates.push("updated_at = CURRENT_TIMESTAMP")
        values.push(existing.id)

        const stmt = this.db.prepare(`UPDATE project_equipment SET ${updates.join(", ")} WHERE id = ?`)
        stmt.run(...values)
      }

      return this.db
        .prepare(`
          SELECT pe.*, em.category, em.equipment_name, em.unit, em.rate
          FROM project_equipment pe
          JOIN equipment_master em ON pe.equipment_id = em.id
          WHERE pe.id = ?
        `)
        .get(existing.id) as ProjectEquipment
    } else {
      // Insert new record
      const stmt = this.db.prepare(`
        INSERT INTO project_equipment (project_id, equipment_id, is_selected, mob_date, demob_date, quantity, hours_month, depreciation_percent, shifts, hire_charges, purchase_value, cost, remarks)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `)
      const result = stmt.run(
        projectId,
        equipmentId,
        data.is_selected || false,
        data.mob_date,
        data.demob_date,
        data.quantity || 1,
        data.hours_month,
        data.depreciation_percent,
        data.shifts,
        data.hire_charges,
        data.purchase_value,
        data.cost,
        data.remarks,
      )

      return this.db
        .prepare(`
          SELECT pe.*, em.category, em.equipment_name, em.unit, em.rate
          FROM project_equipment pe
          JOIN equipment_master em ON pe.equipment_id = em.id
          WHERE pe.id = ?
        `)
        .get(result.lastInsertRowid as number) as ProjectEquipment
    }
  }

  // Project Electricals operations
  getProjectElectricals(projectId: number): ProjectElectricals | undefined {
    return this.db.prepare("SELECT * FROM project_electricals WHERE project_id = ?").get(projectId) as
      | ProjectElectricals
      | undefined
  }

  upsertProjectElectricals(projectId: number, data: Partial<ProjectElectricals>): ProjectElectricals {
    const existing = this.getProjectElectricals(projectId)

    if (existing) {
      const updates: string[] = []
      const values: any[] = []

      Object.entries(data).forEach(([key, value]) => {
        if (
          value !== undefined &&
          key !== "id" &&
          key !== "project_id" &&
          key !== "created_at" &&
          key !== "updated_at"
        ) {
          updates.push(`${key} = ?`)
          values.push(value)
        }
      })

      if (updates.length > 0) {
        updates.push("updated_at = CURRENT_TIMESTAMP")
        values.push(existing.id)

        const stmt = this.db.prepare(`UPDATE project_electricals SET ${updates.join(", ")} WHERE id = ?`)
        stmt.run(...values)
      }

      return this.getProjectElectricals(projectId)!
    } else {
      const stmt = this.db.prepare(`
        INSERT INTO project_electricals (project_id, installation_percent, hv_cables, lv_cables, control_cables, earthing, lighting, others)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
      `)
      stmt.run(
        projectId,
        data.installation_percent || 0,
        data.hv_cables || 0,
        data.lv_cables || 0,
        data.control_cables || 0,
        data.earthing || 0,
        data.lighting || 0,
        data.others || 0,
      )

      return this.getProjectElectricals(projectId)!
    }
  }

  // Project Pipeline Expenses operations
  getProjectPipelineExpenses(projectId: number): ProjectPipelineExpense[] {
    return this.db
      .prepare("SELECT * FROM project_pipeline_expenses WHERE project_id = ? ORDER BY category")
      .all(projectId) as ProjectPipelineExpense[]
  }

  // Project Mechanic Costs operations
  getProjectMechanicCosts(projectId: number): ProjectMechanicCost[] {
    return this.db
      .prepare("SELECT * FROM project_mechanic_costs WHERE project_id = ? ORDER BY category")
      .all(projectId) as ProjectMechanicCost[]
  }

  // Project Misc Expenses operations
  getProjectMiscExpenses(projectId: number): ProjectMiscExpense[] {
    return this.db
      .prepare("SELECT * FROM project_misc_expenses WHERE project_id = ? ORDER BY category")
      .all(projectId) as ProjectMiscExpense[]
  }

  // Project Staff Salary operations
  getProjectStaffSalary(projectId: number): ProjectStaffSalary[] {
    return this.db
      .prepare("SELECT * FROM project_staff_salary WHERE project_id = ? ORDER BY category")
      .all(projectId) as ProjectStaffSalary[]
  }
}
