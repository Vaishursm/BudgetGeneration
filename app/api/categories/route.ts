import { NextResponse } from "next/server"
import { DatabaseOperations } from "@/lib/database"

const db = new DatabaseOperations()

// GET /api/categories - Get all categories
export async function GET() {
  try {
    const categories = db.getAllCategories()
    return NextResponse.json({ categories })
  } catch (error) {
    console.error("Get categories error:", error)
    return NextResponse.json({ error: "Internal server error" }, { status: 500 })
  }
}
