import { type NextRequest, NextResponse } from "next/server"
import { getDatabase } from "@/lib/database"
import bcrypt from "bcryptjs"

interface Project {
  id: number
  project_name: string
  project_description: string
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

export async function POST(request: NextRequest) {
  try {
    const { projectId, password } = await request.json()

    if (!projectId || !password) {
      return NextResponse.json({ message: "Missing project ID or password" }, { status: 400 })
    }

    const db = getDatabase()
    const project = db
      .prepare(`
      SELECT * FROM projects WHERE id = ?
    `)
      .get(projectId) as Project | undefined

    if (!project) {
      return NextResponse.json({ message: "Project not found" }, { status: 404 })
    }

    const isValidPassword = await bcrypt.compare(password, project.password_hash)

    if (!isValidPassword) {
      return NextResponse.json({ message: "Invalid password" }, { status: 401 })
    }

    // Return project data without password hash
    const { password_hash, ...projectData } = project
    return NextResponse.json(projectData)
  } catch (error) {
    console.error("Database error:", error)
    return NextResponse.json({ message: "Failed to verify password" }, { status: 500 })
  }
}
