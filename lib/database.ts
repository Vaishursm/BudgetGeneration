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
}
