"use client"

import { useState, useEffect } from "react"
import { Button } from "@/components/ui/button"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Textarea } from "@/components/ui/textarea"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog"
import { useToast } from "@/hooks/use-toast"
import { FolderOpen, Plus, Save } from "lucide-react"

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
}

export function ProjectDetailsForm() {
  const [projects, setProjects] = useState<Project[]>([])
  const [selectedProject, setSelectedProject] = useState<string>("")
  const [isNewProject, setIsNewProject] = useState(true)
  const [showPasswordDialog, setShowPasswordDialog] = useState(false)
  const [password, setPassword] = useState("")
  const [confirmPassword, setConfirmPassword] = useState("")
  const [isLoading, setIsLoading] = useState(false)
  const { toast } = useToast()

  const [formData, setFormData] = useState({
    project_name: "",
    project_description: "",
    client_name: "",
    project_location: "",
    start_date: "",
    end_date: "",
    workbook_name: "",
    workbook_location: "",
  })

  // Load existing projects on component mount
  useEffect(() => {
    loadProjects()
  }, [])

  const loadProjects = async () => {
    try {
      const response = await fetch("/api/projects")
      if (response.ok) {
        const data = await response.json()
        setProjects(data)
      }
    } catch (error) {
      console.error("Failed to load projects:", error)
    }
  }

  const handleProjectSelection = async (projectId: string) => {
    if (projectId === "new") {
      setIsNewProject(true)
      setSelectedProject("")
      setFormData({
        project_name: "",
        project_description: "",
        client_name: "",
        project_location: "",
        start_date: "",
        end_date: "",
        workbook_name: "",
        workbook_location: "",
      })
    } else {
      setIsNewProject(false)
      setSelectedProject(projectId)
      // Show password dialog for existing project
      setShowPasswordDialog(true)
    }
  }

  const handlePasswordSubmit = async () => {
    if (!selectedProject) return

    try {
      setIsLoading(true)
      const response = await fetch("/api/projects/verify", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ projectId: selectedProject, password }),
      })

      if (response.ok) {
        const project = await response.json()
        setFormData({
          project_name: project.project_name,
          project_description: project.project_description,
          client_name: project.client_name,
          project_location: project.project_location,
          start_date: project.start_date,
          end_date: project.end_date,
          workbook_name: project.workbook_name,
          workbook_location: project.workbook_location,
        })
        setShowPasswordDialog(false)
        setPassword("")
        toast({
          title: "Project Loaded",
          description: "Project details have been loaded successfully.",
        })
      } else {
        toast({
          title: "Invalid Password",
          description: "The password you entered is incorrect.",
          variant: "destructive",
        })
      }
    } catch (error) {
      toast({
        title: "Error",
        description: "Failed to verify password.",
        variant: "destructive",
      })
    } finally {
      setIsLoading(false)
    }
  }

  const handleSaveAndProceed = () => {
    // Validate form
    if (!formData.project_name || !formData.client_name || !formData.start_date || !formData.end_date) {
      toast({
        title: "Validation Error",
        description: "Please fill in all required fields.",
        variant: "destructive",
      })
      return
    }

    // Validate dates
    const startDate = new Date(formData.start_date)
    const endDate = new Date(formData.end_date)
    const today = new Date()
    today.setHours(0, 0, 0, 0)

    if (startDate < today) {
      toast({
        title: "Invalid Date",
        description: "Start date must be a future date.",
        variant: "destructive",
      })
      return
    }

    if (endDate <= startDate) {
      toast({
        title: "Invalid Date",
        description: "End date must be after start date.",
        variant: "destructive",
      })
      return
    }

    setShowPasswordDialog(true)
  }

  const handleCreateProject = async () => {
    if (isNewProject && password !== confirmPassword) {
      toast({
        title: "Password Mismatch",
        description: "Passwords do not match.",
        variant: "destructive",
      })
      return
    }

    try {
      setIsLoading(true)
      const endpoint = isNewProject ? "/api/projects" : `/api/projects/${selectedProject}`
      const method = isNewProject ? "POST" : "PUT"

      const response = await fetch(endpoint, {
        method,
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ ...formData, password }),
      })

      if (response.ok) {
        toast({
          title: "Success",
          description: `Project ${isNewProject ? "created" : "updated"} successfully!`,
        })
        setShowPasswordDialog(false)
        setPassword("")
        setConfirmPassword("")
        loadProjects()
        // Here you would navigate to the main interface screen
      } else {
        const error = await response.json()
        toast({
          title: "Error",
          description: error.message || "Failed to save project.",
          variant: "destructive",
        })
      }
    } catch (error) {
      toast({
        title: "Error",
        description: "Failed to save project.",
        variant: "destructive",
      })
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="min-h-screen bg-background p-6">
      <div className="max-w-4xl mx-auto">
        <Card>
          <CardHeader>
            <CardTitle className="text-2xl font-bold text-center">
              Budget Generation Software - Project Details
            </CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            {/* Project Selection */}
            <div className="space-y-2">
              <Label htmlFor="project-selection">Project Selection</Label>
              <Select onValueChange={handleProjectSelection}>
                <SelectTrigger>
                  <SelectValue placeholder="Select New Project or Existing Project" />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="new">
                    <div className="flex items-center gap-2">
                      <Plus className="h-4 w-4" />
                      New Project
                    </div>
                  </SelectItem>
                  {projects.map((project) => (
                    <SelectItem key={project.id} value={project.id.toString()}>
                      <div className="flex items-center gap-2">
                        <FolderOpen className="h-4 w-4" />
                        {project.project_name}
                      </div>
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            {/* Project Details Form */}
            <div className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="project_name">Project Name *</Label>
                <Input
                  id="project_name"
                  value={formData.project_name}
                  onChange={(e) => setFormData({ ...formData, project_name: e.target.value })}
                  placeholder="Enter project name"
                  disabled={!isNewProject && selectedProject !== ""}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="client_name">Client Name *</Label>
                <Input
                  id="client_name"
                  value={formData.client_name}
                  onChange={(e) => setFormData({ ...formData, client_name: e.target.value })}
                  placeholder="Enter client name"
                  disabled={!isNewProject && selectedProject !== ""}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="project_description">Project Description</Label>
                <Textarea
                  id="project_description"
                  value={formData.project_description}
                  onChange={(e) => setFormData({ ...formData, project_description: e.target.value })}
                  placeholder="Enter project description"
                  disabled={!isNewProject && selectedProject !== ""}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="project_location">Project Location *</Label>
                <Input
                  id="project_location"
                  value={formData.project_location}
                  onChange={(e) => setFormData({ ...formData, project_location: e.target.value })}
                  placeholder="Enter project location"
                  disabled={!isNewProject && selectedProject !== ""}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="start_date">Start Date *</Label>
                <Input
                  id="start_date"
                  type="date"
                  value={formData.start_date}
                  onChange={(e) => setFormData({ ...formData, start_date: e.target.value })}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="end_date">End Date *</Label>
                <Input
                  id="end_date"
                  type="date"
                  value={formData.end_date}
                  onChange={(e) => setFormData({ ...formData, end_date: e.target.value })}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="workbook_name">Workbook Name</Label>
                <Input
                  id="workbook_name"
                  value={formData.workbook_name}
                  onChange={(e) => setFormData({ ...formData, workbook_name: e.target.value })}
                  placeholder="Enter Excel workbook name"
                  disabled={!isNewProject && selectedProject !== ""}
                />
              </div>

              <div className="space-y-2">
                <Label htmlFor="workbook_location">Workbook Save Location</Label>
                <Input
                  id="workbook_location"
                  value={formData.workbook_location}
                  onChange={(e) => setFormData({ ...formData, workbook_location: e.target.value })}
                  placeholder="Enter save location path"
                  disabled={!isNewProject && selectedProject !== ""}
                />
              </div>
            </div>

            {/* Save & Proceed Button */}
            <div className="flex justify-center pt-4">
              <Button onClick={handleSaveAndProceed} size="lg" className="px-8">
                <Save className="mr-2 h-4 w-4" />
                Save & Proceed
              </Button>
            </div>
          </CardContent>
        </Card>

        {/* Password Dialog */}
        <Dialog open={showPasswordDialog} onOpenChange={setShowPasswordDialog}>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>{isNewProject ? "Create Project Password" : "Enter Project Password"}</DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              <div className="space-y-2">
                <Label htmlFor="password">Password</Label>
                <Input
                  id="password"
                  type="password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="Enter password"
                />
              </div>
              {isNewProject && (
                <div className="space-y-2">
                  <Label htmlFor="confirm_password">Confirm Password</Label>
                  <Input
                    id="confirm_password"
                    type="password"
                    value={confirmPassword}
                    onChange={(e) => setConfirmPassword(e.target.value)}
                    placeholder="Confirm password"
                  />
                </div>
              )}
              <div className="flex justify-end gap-2">
                <Button variant="outline" onClick={() => setShowPasswordDialog(false)}>
                  Cancel
                </Button>
                <Button onClick={isNewProject ? handleCreateProject : handlePasswordSubmit} disabled={isLoading}>
                  {isLoading ? "Processing..." : isNewProject ? "Create Project" : "Load Project"}
                </Button>
              </div>
            </div>
          </DialogContent>
        </Dialog>
      </div>
    </div>
  )
}
