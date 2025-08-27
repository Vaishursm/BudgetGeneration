import { type NextRequest, NextResponse } from "next/server"
import { getDatabase } from "@/lib/database"
import bcrypt from "bcryptjs"

export async function GET() {
  try {
    const db = getDatabase()
    const projects = db
      .prepare(`
      SELECT id, project_name, client_name, project_location, start_date, end_date, created_at
      FROM projects 
      ORDER BY created_at DESC
    `)
      .all()

    return NextResponse.json(projects)
  } catch (error) {
    console.error("Database error:", error)
    return NextResponse.json({ message: "Failed to fetch projects" }, { status: 500 })
  }
}

export async function POST(request: NextRequest) {
  try {
    const body = await request.json()
    const {
      project_name,
      project_description,
      client_name,
      project_location,
      start_date,
      end_date,
      workbook_name,
      workbook_location,
      password,
    } = body

    // Validate required fields
    if (!project_name || !client_name || !start_date || !end_date || !password) {
      return NextResponse.json({ message: "Missing required fields" }, { status: 400 })
    }

    // Hash password
    const password_hash = await bcrypt.hash(password, 10)

    const db = getDatabase()
    const result = db
      .prepare(`
      INSERT INTO projects (
        project_name, project_description, client_name, project_location,
        start_date, end_date, workbook_name, workbook_location, password_hash
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `)
      .run(
        project_name,
        project_description || "",
        client_name,
        project_location,
        start_date,
        end_date,
        workbook_name || "",
        workbook_location || "",
        password_hash,
      )

    return NextResponse.json({
      message: "Project created successfully",
      projectId: result.lastInsertRowid,
    })
  } catch (error) {
    console.error("Database error:", error)
    return NextResponse.json({ message: "Failed to create project" }, { status: 500 })
  }
}
