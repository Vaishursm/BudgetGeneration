import bcrypt from "bcryptjs"
import jwt from "jsonwebtoken"
import { DatabaseOperations } from "./database"

const JWT_SECRET = process.env.JWT_SECRET || "your-secret-key-change-in-production"
const db = new DatabaseOperations()

export interface AuthUser {
  id: number
  email: string
  username: string
  first_name?: string
  last_name?: string
}

export class AuthService {
  static async hashPassword(password: string): Promise<string> {
    return bcrypt.hash(password, 10)
  }

  static async verifyPassword(password: string, hash: string): Promise<boolean> {
    return bcrypt.compare(password, hash)
  }

  static generateToken(user: AuthUser): string {
    return jwt.sign(
      {
        id: user.id,
        email: user.email,
        username: user.username,
      },
      JWT_SECRET,
      { expiresIn: "7d" },
    )
  }

  static verifyToken(token: string): AuthUser | null {
    try {
      return jwt.verify(token, JWT_SECRET) as AuthUser
    } catch {
      return null
    }
  }

  static async login(email: string, password: string): Promise<{ user: AuthUser; token: string } | null> {
    const user = db.getUserByEmail(email)
    if (!user) return null

    const isValidPassword = await this.verifyPassword(password, user.password_hash)
    if (!isValidPassword) return null

    const authUser: AuthUser = {
      id: user.id,
      email: user.email,
      username: user.username,
      first_name: user.first_name,
      last_name: user.last_name,
    }

    const token = this.generateToken(authUser)
    return { user: authUser, token }
  }

  static async register(userData: {
    email: string
    username: string
    password: string
    first_name?: string
    last_name?: string
  }): Promise<{ user: AuthUser; token: string } | null> {
    try {
      const hashedPassword = await this.hashPassword(userData.password)

      const user = db.createUser({
        email: userData.email,
        username: userData.username,
        password_hash: hashedPassword,
        first_name: userData.first_name,
        last_name: userData.last_name,
      })

      const authUser: AuthUser = {
        id: user.id,
        email: user.email,
        username: user.username,
        first_name: user.first_name,
        last_name: user.last_name,
      }

      const token = this.generateToken(authUser)
      return { user: authUser, token }
    } catch {
      return null
    }
  }
}
