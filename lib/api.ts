// API client for frontend-backend communication
const API_BASE = "/api"

export interface User {
  id: number
  email: string
  username: string
  first_name?: string
  last_name?: string
}

export interface Post {
  id: number
  title: string
  content: string
  author_id: number
  status: "draft" | "published" | "archived"
  created_at: string
  updated_at: string
  username?: string
  first_name?: string
  last_name?: string
  categories?: Category[]
}

export interface Category {
  id: number
  name: string
  description?: string
  created_at: string
}

export interface AuthResponse {
  user: User
  token: string
  message: string
}

class ApiClient {
  private getAuthHeaders() {
    const token = localStorage.getItem("auth_token")
    return token ? { Authorization: `Bearer ${token}` } : {}
  }

  async login(email: string, password: string): Promise<AuthResponse> {
    const response = await fetch(`${API_BASE}/auth/login`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || "Login failed")
    }

    return response.json()
  }

  async register(userData: {
    email: string
    username: string
    password: string
    first_name?: string
    last_name?: string
  }): Promise<AuthResponse> {
    const response = await fetch(`${API_BASE}/auth/register`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(userData),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || "Registration failed")
    }

    return response.json()
  }

  async getPosts(status?: string): Promise<{ posts: Post[] }> {
    const url = status ? `${API_BASE}/posts?status=${status}` : `${API_BASE}/posts`
    const response = await fetch(url)

    if (!response.ok) {
      throw new Error("Failed to fetch posts")
    }

    return response.json()
  }

  async getPost(id: number): Promise<{ post: Post }> {
    const response = await fetch(`${API_BASE}/posts/${id}`)

    if (!response.ok) {
      throw new Error("Failed to fetch post")
    }

    return response.json()
  }

  async createPost(postData: {
    title: string
    content: string
    status?: string
  }): Promise<{ post: Post; message: string }> {
    const response = await fetch(`${API_BASE}/posts`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        ...this.getAuthHeaders(),
      },
      body: JSON.stringify(postData),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || "Failed to create post")
    }

    return response.json()
  }

  async updatePost(id: number, postData: Partial<Post>): Promise<{ post: Post; message: string }> {
    const response = await fetch(`${API_BASE}/posts/${id}`, {
      method: "PUT",
      headers: {
        "Content-Type": "application/json",
        ...this.getAuthHeaders(),
      },
      body: JSON.stringify(postData),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || "Failed to update post")
    }

    return response.json()
  }

  async deletePost(id: number): Promise<{ message: string }> {
    const response = await fetch(`${API_BASE}/posts/${id}`, {
      method: "DELETE",
      headers: this.getAuthHeaders(),
    })

    if (!response.ok) {
      const error = await response.json()
      throw new Error(error.error || "Failed to delete post")
    }

    return response.json()
  }

  async getCategories(): Promise<{ categories: Category[] }> {
    const response = await fetch(`${API_BASE}/categories`)

    if (!response.ok) {
      throw new Error("Failed to fetch categories")
    }

    return response.json()
  }
}

export const api = new ApiClient()
