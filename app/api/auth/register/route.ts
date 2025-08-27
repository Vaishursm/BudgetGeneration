import { type NextRequest, NextResponse } from "next/server"
import { AuthService } from "@/lib/auth"

export async function POST(request: NextRequest) {
  try {
    const { email, username, password, first_name, last_name } = await request.json()

    if (!email || !username || !password) {
      return NextResponse.json({ error: "Email, username, and password are required" }, { status: 400 })
    }

    const result = await AuthService.register({
      email,
      username,
      password,
      first_name,
      last_name,
    })

    if (!result) {
      return NextResponse.json({ error: "User already exists or registration failed" }, { status: 409 })
    }

    return NextResponse.json(
      {
        message: "Registration successful",
        user: result.user,
        token: result.token,
      },
      { status: 201 },
    )
  } catch (error) {
    console.error("Registration error:", error)
    return NextResponse.json({ error: "Internal server error" }, { status: 500 })
  }
}
