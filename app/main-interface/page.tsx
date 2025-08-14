"use client"

import { useState } from "react"
import { Button } from "@/components/ui/button"
import { Checkbox } from "@/components/ui/checkbox"
import { RadioGroup, RadioGroupItem } from "@/components/ui/radio-group"
import { Label } from "@/components/ui/label"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs"
import { Input } from "@/components/ui/input"
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card"
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select"
import { useToast } from "@/hooks/use-toast"

const equipmentCategories = [
  "Major Concrete",
  "Major Conveyance",
  "Major Crane",
  "Major DG Sets",
  "Major Material Handling",
  "Major Non-Concrete",
  "Major Others",
  "Minor E Equipments",
  "Hired Equipments",
  "Fixed Exp. - Tower Crane",
  "Fixed Exp. - BP Related",
  "Lighting/Single Phase Equips",
]

const sampleEquipmentData = {
  "Major Concrete": [
    {
      id: 1,
      name: "Concrete Mixer 10/7",
      unit: "No",
      rate: 2500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 200,
      depreciation: 10,
      shifts: 2,
    },
    {
      id: 2,
      name: "Concrete Pump",
      unit: "No",
      rate: 3500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      depreciation: 12,
      shifts: 2,
    },
    {
      id: 3,
      name: "Vibrator Needle",
      unit: "No",
      rate: 150,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 4,
      hrsMonth: 160,
      depreciation: 15,
      shifts: 1,
    },
    {
      id: 4,
      name: "Transit Mixer",
      unit: "No",
      rate: 4500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 220,
      depreciation: 8,
      shifts: 2,
    },
  ],
  "Major Conveyance": [
    {
      id: 5,
      name: "Belt Conveyor",
      unit: "No",
      rate: 1800,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 200,
      depreciation: 10,
      shifts: 2,
    },
    {
      id: 6,
      name: "Bucket Elevator",
      unit: "No",
      rate: 2200,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      depreciation: 12,
      shifts: 1,
    },
    {
      id: 7,
      name: "Screw Conveyor",
      unit: "No",
      rate: 1500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 160,
      depreciation: 15,
      shifts: 1,
    },
  ],
  "Major Crane": [
    {
      id: 8,
      name: "Tower Crane 6T",
      unit: "No",
      rate: 8500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 200,
      depreciation: 8,
      shifts: 2,
    },
    {
      id: 9,
      name: "Mobile Crane 25T",
      unit: "No",
      rate: 6500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      depreciation: 10,
      shifts: 2,
    },
    {
      id: 10,
      name: "Crawler Crane 40T",
      unit: "No",
      rate: 9500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 220,
      depreciation: 8,
      shifts: 2,
    },
  ],
  "Major DG Sets": [
    {
      id: 11,
      name: "DG Set 125 KVA",
      unit: "No",
      rate: 3500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 200,
      depreciation: 12,
      shifts: 2,
    },
    {
      id: 12,
      name: "DG Set 250 KVA",
      unit: "No",
      rate: 5500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      depreciation: 10,
      shifts: 2,
    },
    {
      id: 13,
      name: "DG Set 500 KVA",
      unit: "No",
      rate: 8500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 220,
      depreciation: 8,
      shifts: 2,
    },
  ],
  "Major Material Handling": [
    {
      id: 14,
      name: "Forklift 3T",
      unit: "No",
      rate: 2800,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 200,
      depreciation: 12,
      shifts: 2,
    },
    {
      id: 15,
      name: "Reach Stacker",
      unit: "No",
      rate: 4500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      depreciation: 10,
      shifts: 2,
    },
    {
      id: 16,
      name: "Telehandler",
      unit: "No",
      rate: 3200,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 160,
      depreciation: 15,
      shifts: 1,
    },
  ],
  "Major Non-Concrete": [
    {
      id: 17,
      name: "Excavator 20T",
      unit: "No",
      rate: 4500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 200,
      depreciation: 10,
      shifts: 2,
    },
    {
      id: 18,
      name: "Bulldozer D6",
      unit: "No",
      rate: 5500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      depreciation: 8,
      shifts: 2,
    },
    {
      id: 19,
      name: "Grader 140H",
      unit: "No",
      rate: 6500,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 220,
      depreciation: 8,
      shifts: 2,
    },
  ],
  "Major Others": [
    {
      id: 20,
      name: "Compressor 750 CFM",
      unit: "No",
      rate: 2200,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 200,
      depreciation: 12,
      shifts: 2,
    },
    {
      id: 21,
      name: "Welding Machine 400A",
      unit: "No",
      rate: 800,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 3,
      hrsMonth: 160,
      depreciation: 15,
      shifts: 1,
    },
    {
      id: 22,
      name: 'Water Pump 6"',
      unit: "No",
      rate: 450,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 4,
      hrsMonth: 180,
      depreciation: 15,
      shifts: 1,
    },
  ],
  "Minor E Equipments": [
    {
      id: 23,
      name: "Hand Drill",
      unit: "No",
      rate: 120,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 5,
      hrsMonth: 100,
      pVal: 0,
    },
    {
      id: 24,
      name: "Angle Grinder",
      unit: "No",
      rate: 80,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 8,
      hrsMonth: 120,
      pVal: 0,
    },
    {
      id: 25,
      name: "Circular Saw",
      unit: "No",
      rate: 200,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 3,
      hrsMonth: 140,
      pVal: 0,
    },
  ],
  "Hired Equipments": [
    {
      id: 26,
      name: "Hired Excavator",
      unit: "No",
      rate: 0,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 200,
      hireCharges: 4500,
    },
    {
      id: 27,
      name: "Hired Crane",
      unit: "No",
      rate: 0,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 1,
      hrsMonth: 180,
      hireCharges: 6500,
    },
    {
      id: 28,
      name: "Hired Truck",
      unit: "No",
      rate: 0,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 2,
      hrsMonth: 220,
      hireCharges: 2500,
    },
  ],
  "Fixed Exp. - Tower Crane": [
    {
      id: 29,
      name: "Tower Crane Foundation",
      unit: "LS",
      rate: 0,
      selected: false,
      qty: 1,
      cost: 150000,
      remarks: "Foundation work",
    },
    {
      id: 30,
      name: "Tower Crane Erection",
      unit: "LS",
      rate: 0,
      selected: false,
      qty: 1,
      cost: 80000,
      remarks: "Erection charges",
    },
    {
      id: 31,
      name: "Tower Crane Dismantling",
      unit: "LS",
      rate: 0,
      selected: false,
      qty: 1,
      cost: 60000,
      remarks: "Dismantling charges",
    },
  ],
  "Fixed Exp. - BP Related": [
    {
      id: 32,
      name: "Site Office Setup",
      unit: "LS",
      rate: 0,
      selected: false,
      qty: 1,
      cost: 200000,
      remarks: "Office setup",
    },
    {
      id: 33,
      name: "Labour Camp",
      unit: "LS",
      rate: 0,
      selected: false,
      qty: 1,
      cost: 300000,
      remarks: "Camp construction",
    },
    {
      id: 34,
      name: "Boundary Wall",
      unit: "LS",
      rate: 0,
      selected: false,
      qty: 1,
      cost: 150000,
      remarks: "Security wall",
    },
  ],
  "Lighting/Single Phase Equips": [
    {
      id: 35,
      name: "LED Flood Light 100W",
      unit: "No",
      rate: 0,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 20,
    },
    {
      id: 36,
      name: "Street Light 50W",
      unit: "No",
      rate: 0,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 15,
    },
    {
      id: 37,
      name: "Emergency Light",
      unit: "No",
      rate: 0,
      selected: false,
      mobDate: "2024-02-01",
      demobDate: "2024-06-30",
      qty: 10,
    },
  ],
}

const electricalData = {
  installationPercent: 15,
  breakup: {
    cabling: 40,
    switchgear: 25,
    lighting: 20,
    others: 15,
  },
}

const pipelineExpenses = [
  { id: 1, category: "Water Supply", qty: 1, costPerUnit: 50000, amount: 50000, remarks: "Main water line" },
  { id: 2, category: "Sewerage", qty: 1, costPerUnit: 75000, amount: 75000, remarks: "Sewerage connection" },
  {
    id: 3,
    category: "Electrical Connection",
    qty: 1,
    costPerUnit: 100000,
    amount: 100000,
    remarks: "Main electrical supply",
  },
  { id: 4, category: "Telecom", qty: 1, costPerUnit: 25000, amount: 25000, remarks: "Internet & phone lines" },
]

const electMechanicCost = [
  { id: 1, category: "Electrician", nos: 3, salaryPerMonth: 25000, noOfMonths: 12, salaryCost: 900000 },
  { id: 2, category: "Mechanic", nos: 2, salaryPerMonth: 22000, noOfMonths: 12, salaryCost: 528000 },
  { id: 3, category: "Welder", nos: 4, salaryPerMonth: 20000, noOfMonths: 10, salaryCost: 800000 },
  { id: 4, category: "Helper", nos: 6, salaryPerMonth: 15000, noOfMonths: 12, salaryCost: 1080000 },
]

const miscExpenses = [
  { id: 1, type: "Insurance", amount: 150000, remarks: "Equipment insurance" },
  { id: 2, type: "Transportation", amount: 200000, remarks: "Equipment transportation" },
  { id: 3, type: "Permits & Licenses", amount: 75000, remarks: "Various permits" },
  { id: 4, type: "Safety Equipment", amount: 100000, remarks: "Safety gear and equipment" },
]

const staffSalary = [
  { id: 1, category: "Project Manager", nos: 1, salaryPerMonth: 80000, noOfMonths: 12, salaryCost: 960000 },
  { id: 2, category: "Site Engineer", nos: 2, salaryPerMonth: 45000, noOfMonths: 12, salaryCost: 1080000 },
  { id: 3, category: "Supervisor", nos: 3, salaryPerMonth: 30000, noOfMonths: 12, salaryCost: 1080000 },
  { id: 4, category: "Safety Officer", nos: 1, salaryPerMonth: 35000, noOfMonths: 12, salaryCost: 420000 },
]

export default function MainInterface() {
  const [selectedCategory, setSelectedCategory] = useState("Major Concrete")
  const [equipment, setEquipment] = useState(sampleEquipmentData)
  const [activeTab, setActiveTab] = useState("equipment")
  const [electrical, setElectrical] = useState(electricalData)
  const [pipeline, setPipeline] = useState(pipelineExpenses)
  const [electMechanic, setElectMechanic] = useState(electMechanicCost)
  const [misc, setMisc] = useState(miscExpenses)
  const [staff, setStaff] = useState(staffSalary)
  const { toast } = useToast()

  const currentEquipment = equipment[selectedCategory] || []

  const handleSelectAll = () => {
    setEquipment((prev) => ({
      ...prev,
      [selectedCategory]: prev[selectedCategory]?.map((item) => ({ ...item, selected: true })) || [],
    }))
    toast({ title: "All items selected", description: `Selected all equipment in ${selectedCategory}` })
  }

  const handleUnselectAll = () => {
    setEquipment((prev) => ({
      ...prev,
      [selectedCategory]: prev[selectedCategory]?.map((item) => ({ ...item, selected: false })) || [],
    }))
    toast({ title: "All items unselected", description: `Unselected all equipment in ${selectedCategory}` })
  }

  const handleEquipmentToggle = (id: number) => {
    setEquipment((prev) => ({
      ...prev,
      [selectedCategory]:
        prev[selectedCategory]?.map((item) => (item.id === id ? { ...item, selected: !item.selected } : item)) || [],
    }))
  }

  const handleEquipmentUpdate = (id: number, field: string, value: any) => {
    setEquipment((prev) => ({
      ...prev,
      [selectedCategory]:
        prev[selectedCategory]?.map((item) => (item.id === id ? { ...item, [field]: value } : item)) || [],
    }))
  }

  const addEquipment = (baseItem: any) => {
    const newItem = {
      ...baseItem,
      id: Date.now(),
      selected: false,
    }
    setEquipment((prev) => ({
      ...prev,
      [selectedCategory]: [...(prev[selectedCategory] || []), newItem],
    }))
    toast({ title: "Equipment added", description: `Added new ${baseItem.name}` })
  }

  const getSelectedCount = () => {
    return currentEquipment.filter((item) => item.selected).length
  }

  const viewSelectedItems = () => {
    const selected = currentEquipment.filter((item) => item.selected)
    toast({
      title: "Selected Items",
      description: `${selected.length} items selected in ${selectedCategory}`,
    })
  }

  const validateDate = (date: string) => {
    const today = new Date().toISOString().split("T")[0]
    return date >= today
  }

  const renderEquipmentTable = () => {
    const isFixedExp = selectedCategory.includes("Fixed Exp.")
    const isMinorEquipment = selectedCategory === "Minor E Equipments"
    const isHiredEquipment = selectedCategory === "Hired Equipments"
    const isLighting = selectedCategory === "Lighting/Single Phase Equips"

    return (
      <div className="overflow-x-auto">
        <table className="w-full border-collapse border">
          <thead>
            <tr className="bg-gray-100">
              <th className="border p-2 text-left">Select</th>
              <th className="border p-2 text-left bg-gray-200">Equipment Name</th>
              <th className="border p-2 text-left bg-gray-200">Unit</th>
              {!isFixedExp && <th className="border p-2 text-left bg-gray-200">Rate</th>}
              {!isLighting && <th className="border p-2 text-left">Mob Date</th>}
              {!isLighting && <th className="border p-2 text-left">Demob Date</th>}
              <th className="border p-2 text-left">Qty</th>
              {!isFixedExp && !isLighting && <th className="border p-2 text-left">Hrs/Month</th>}
              {!isFixedExp && !isMinorEquipment && !isHiredEquipment && !isLighting && (
                <th className="border p-2 text-left">Depreciation %</th>
              )}
              {!isFixedExp && !isMinorEquipment && !isHiredEquipment && !isLighting && (
                <th className="border p-2 text-left">Shifts</th>
              )}
              {isMinorEquipment && <th className="border p-2 text-left">P.Val</th>}
              {isHiredEquipment && <th className="border p-2 text-left">Hire Charges</th>}
              {isFixedExp && <th className="border p-2 text-left">Cost</th>}
              {isFixedExp && <th className="border p-2 text-left">Remarks</th>}
              <th className="border p-2 text-left">Action</th>
            </tr>
          </thead>
          <tbody>
            {currentEquipment.map((item) => (
              <tr key={item.id} className="hover:bg-gray-50">
                <td className="border p-2">
                  <Checkbox checked={item.selected} onCheckedChange={() => handleEquipmentToggle(item.id)} />
                </td>
                <td className="border p-2 bg-gray-100">{item.name}</td>
                <td className="border p-2 bg-gray-100">{item.unit}</td>
                {!isFixedExp && <td className="border p-2 bg-gray-100">₹{item.rate}</td>}
                {!isLighting && (
                  <td className="border p-2">
                    <Input
                      type="date"
                      value={item.mobDate}
                      onChange={(e) => handleEquipmentUpdate(item.id, "mobDate", e.target.value)}
                      className="w-36"
                    />
                  </td>
                )}
                {!isLighting && (
                  <td className="border p-2">
                    <Input
                      type="date"
                      value={item.demobDate}
                      onChange={(e) => handleEquipmentUpdate(item.id, "demobDate", e.target.value)}
                      className="w-36"
                      min={item.mobDate}
                    />
                  </td>
                )}
                <td className="border p-2">
                  <Input
                    type="number"
                    value={item.qty}
                    onChange={(e) => handleEquipmentUpdate(item.id, "qty", Number.parseInt(e.target.value))}
                    className="w-20"
                    min="1"
                  />
                </td>
                {!isFixedExp && !isLighting && (
                  <td className="border p-2">
                    <Input
                      type="number"
                      value={item.hrsMonth}
                      onChange={(e) => handleEquipmentUpdate(item.id, "hrsMonth", Number.parseInt(e.target.value))}
                      className="w-24"
                      min="1"
                    />
                  </td>
                )}
                {!isFixedExp && !isMinorEquipment && !isHiredEquipment && !isLighting && (
                  <td className="border p-2">
                    <Input
                      type="number"
                      value={item.depreciation}
                      onChange={(e) =>
                        handleEquipmentUpdate(item.id, "depreciation", Number.parseFloat(e.target.value))
                      }
                      className="w-20"
                      min="0"
                      max="100"
                      step="0.1"
                    />
                  </td>
                )}
                {!isFixedExp && !isMinorEquipment && !isHiredEquipment && !isLighting && (
                  <td className="border p-2">
                    <Select
                      value={item.shifts?.toString()}
                      onValueChange={(value) => handleEquipmentUpdate(item.id, "shifts", Number.parseInt(value))}
                    >
                      <SelectTrigger className="w-20">
                        <SelectValue />
                      </SelectTrigger>
                      <SelectContent>
                        <SelectItem value="1">1</SelectItem>
                        <SelectItem value="2">2</SelectItem>
                        <SelectItem value="3">3</SelectItem>
                      </SelectContent>
                    </Select>
                  </td>
                )}
                {isMinorEquipment && (
                  <td className="border p-2">
                    <Input
                      type="number"
                      value={item.pVal}
                      onChange={(e) => handleEquipmentUpdate(item.id, "pVal", Number.parseInt(e.target.value))}
                      className="w-24"
                      min="0"
                    />
                  </td>
                )}
                {isHiredEquipment && (
                  <td className="border p-2">
                    <Input
                      type="number"
                      value={item.hireCharges}
                      onChange={(e) => handleEquipmentUpdate(item.id, "hireCharges", Number.parseInt(e.target.value))}
                      className="w-24"
                      min="0"
                    />
                  </td>
                )}
                {isFixedExp && (
                  <td className="border p-2">
                    <Input
                      type="number"
                      value={item.cost}
                      onChange={(e) => handleEquipmentUpdate(item.id, "cost", Number.parseInt(e.target.value))}
                      className="w-24"
                      min="0"
                    />
                  </td>
                )}
                {isFixedExp && (
                  <td className="border p-2">
                    <Input
                      value={item.remarks}
                      onChange={(e) => handleEquipmentUpdate(item.id, "remarks", e.target.value)}
                      className="w-32"
                    />
                  </td>
                )}
                <td className="border p-2">
                  <Button size="sm" variant="outline" onClick={() => addEquipment(item)}>
                    Add
                  </Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50 p-4">
      <div className="max-w-full mx-auto">
        <Card>
          <CardHeader>
            <CardTitle className="text-2xl font-bold text-center">Budget Generation - Main Interface</CardTitle>
          </CardHeader>
          <CardContent className="space-y-6">
            {/* Tabs */}
            <Tabs value={activeTab} onValueChange={setActiveTab}>
              <TabsList className="grid w-full grid-cols-6">
                <TabsTrigger value="equipment">Equipment Entry</TabsTrigger>
                <TabsTrigger value="electricals">Electricals</TabsTrigger>
                <TabsTrigger value="pipeline">Pipeline Expenses</TabsTrigger>
                <TabsTrigger value="electmechanic">Elect/Mechanic Cost</TabsTrigger>
                <TabsTrigger value="misc">Misc & Non-ERP</TabsTrigger>
                <TabsTrigger value="staff">Staff Salary</TabsTrigger>
              </TabsList>

              <TabsContent value="equipment" className="space-y-6">
                {/* Control Buttons */}
                <div className="flex gap-4 justify-center flex-wrap">
                  <Button onClick={handleSelectAll} variant="outline">
                    Select All
                  </Button>
                  <Button onClick={handleUnselectAll} variant="outline">
                    Unselect All
                  </Button>
                  <Button variant="outline" onClick={viewSelectedItems}>
                    View Selected Items List ({getSelectedCount()})
                  </Button>
                </div>

                {/* Equipment Table */}
                <div className="border rounded-lg overflow-hidden">
                  <div className="bg-blue-100 p-4">
                    <h3 className="font-semibold text-lg">{selectedCategory} Equipment</h3>
                  </div>
                  {renderEquipmentTable()}
                </div>

                {/* Category Radio Buttons */}
                <div className="border-t pt-6">
                  <h4 className="font-semibold mb-4">Equipment Categories:</h4>
                  <RadioGroup
                    value={selectedCategory}
                    onValueChange={setSelectedCategory}
                    className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2"
                  >
                    {equipmentCategories.map((category) => (
                      <div key={category} className="flex items-center space-x-2">
                        <RadioGroupItem value={category} id={category} />
                        <Label htmlFor={category} className="text-sm cursor-pointer">
                          {category}
                        </Label>
                      </div>
                    ))}
                  </RadioGroup>
                </div>
              </TabsContent>

              <TabsContent value="electricals" className="space-y-6">
                <Card>
                  <CardHeader>
                    <CardTitle>Electrical Installation</CardTitle>
                  </CardHeader>
                  <CardContent className="space-y-4">
                    <div>
                      <Label>Electrical Installation % at the top</Label>
                      <Input
                        type="number"
                        value={electrical.installationPercent}
                        onChange={(e) =>
                          setElectrical((prev) => ({ ...prev, installationPercent: Number.parseFloat(e.target.value) }))
                        }
                        className="w-32"
                        min="0"
                        max="100"
                        step="0.1"
                      />
                    </div>
                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <Label>Cabling %</Label>
                        <Input
                          type="number"
                          value={electrical.breakup.cabling}
                          onChange={(e) =>
                            setElectrical((prev) => ({
                              ...prev,
                              breakup: { ...prev.breakup, cabling: Number.parseFloat(e.target.value) },
                            }))
                          }
                          className="w-24"
                        />
                      </div>
                      <div>
                        <Label>Switchgear %</Label>
                        <Input
                          type="number"
                          value={electrical.breakup.switchgear}
                          onChange={(e) =>
                            setElectrical((prev) => ({
                              ...prev,
                              breakup: { ...prev.breakup, switchgear: Number.parseFloat(e.target.value) },
                            }))
                          }
                          className="w-24"
                        />
                      </div>
                      <div>
                        <Label>Lighting %</Label>
                        <Input
                          type="number"
                          value={electrical.breakup.lighting}
                          onChange={(e) =>
                            setElectrical((prev) => ({
                              ...prev,
                              breakup: { ...prev.breakup, lighting: Number.parseFloat(e.target.value) },
                            }))
                          }
                          className="w-24"
                        />
                      </div>
                      <div>
                        <Label>Others %</Label>
                        <Input
                          type="number"
                          value={electrical.breakup.others}
                          onChange={(e) =>
                            setElectrical((prev) => ({
                              ...prev,
                              breakup: { ...prev.breakup, others: Number.parseFloat(e.target.value) },
                            }))
                          }
                          className="w-24"
                        />
                      </div>
                    </div>
                    <div className="flex gap-4 mt-6">
                      <Button variant="outline" onClick={() => setActiveTab("equipment")}>
                        Goto Equipment Entry
                      </Button>
                      <Button variant="outline" onClick={() => setActiveTab("pipeline")}>
                        Goto Pipeline Expenses Entry
                      </Button>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="pipeline" className="space-y-6">
                <Card>
                  <CardHeader>
                    <CardTitle>Pipeline Expenses</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse border">
                        <thead>
                          <tr className="bg-gray-100">
                            <th className="border p-2 text-left">Category</th>
                            <th className="border p-2 text-left">Qty</th>
                            <th className="border p-2 text-left">Cost per Unit</th>
                            <th className="border p-2 text-left bg-gray-200">Amount</th>
                            <th className="border p-2 text-left">Remarks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pipeline.map((item) => (
                            <tr key={item.id}>
                              <td className="border p-2 bg-gray-100">{item.category}</td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.qty}
                                  onChange={(e) => {
                                    const newQty = Number.parseInt(e.target.value)
                                    setPipeline((prev) =>
                                      prev.map((p) =>
                                        p.id === item.id ? { ...p, qty: newQty, amount: newQty * p.costPerUnit } : p,
                                      ),
                                    )
                                  }}
                                  className="w-20"
                                />
                              </td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.costPerUnit}
                                  onChange={(e) => {
                                    const newCost = Number.parseInt(e.target.value)
                                    setPipeline((prev) =>
                                      prev.map((p) =>
                                        p.id === item.id ? { ...p, costPerUnit: newCost, amount: p.qty * newCost } : p,
                                      ),
                                    )
                                  }}
                                  className="w-24"
                                />
                              </td>
                              <td className="border p-2 bg-gray-200">₹{item.amount.toLocaleString()}</td>
                              <td className="border p-2">
                                <Input
                                  value={item.remarks}
                                  onChange={(e) =>
                                    setPipeline((prev) =>
                                      prev.map((p) => (p.id === item.id ? { ...p, remarks: e.target.value } : p)),
                                    )
                                  }
                                  className="w-40"
                                />
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="electmechanic" className="space-y-6">
                <Card>
                  <CardHeader>
                    <CardTitle>Elect/Mechanic Cost</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse border">
                        <thead>
                          <tr className="bg-gray-100">
                            <th className="border p-2 text-left">Category</th>
                            <th className="border p-2 text-left">Nos</th>
                            <th className="border p-2 text-left">Salary per Month</th>
                            <th className="border p-2 text-left">No. of Months</th>
                            <th className="border p-2 text-left bg-gray-200">Salary Cost</th>
                          </tr>
                        </thead>
                        <tbody>
                          {electMechanic.map((item) => (
                            <tr key={item.id}>
                              <td className="border p-2 bg-gray-100">{item.category}</td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.nos}
                                  onChange={(e) => {
                                    const newNos = Number.parseInt(e.target.value)
                                    setElectMechanic((prev) =>
                                      prev.map((em) =>
                                        em.id === item.id
                                          ? {
                                              ...em,
                                              nos: newNos,
                                              salaryCost: newNos * em.salaryPerMonth * em.noOfMonths,
                                            }
                                          : em,
                                      ),
                                    )
                                  }}
                                  className="w-20"
                                />
                              </td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.salaryPerMonth}
                                  onChange={(e) => {
                                    const newSalary = Number.parseInt(e.target.value)
                                    setElectMechanic((prev) =>
                                      prev.map((em) =>
                                        em.id === item.id
                                          ? {
                                              ...em,
                                              salaryPerMonth: newSalary,
                                              salaryCost: em.nos * newSalary * em.noOfMonths,
                                            }
                                          : em,
                                      ),
                                    )
                                  }}
                                  className="w-24"
                                />
                              </td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.noOfMonths}
                                  onChange={(e) => {
                                    const newMonths = Number.parseInt(e.target.value)
                                    setElectMechanic((prev) =>
                                      prev.map((em) =>
                                        em.id === item.id
                                          ? {
                                              ...em,
                                              noOfMonths: newMonths,
                                              salaryCost: em.nos * em.salaryPerMonth * newMonths,
                                            }
                                          : em,
                                      ),
                                    )
                                  }}
                                  className="w-20"
                                  max="12"
                                />
                              </td>
                              <td className="border p-2 bg-gray-200">₹{item.salaryCost.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="misc" className="space-y-6">
                <Card>
                  <CardHeader>
                    <CardTitle>Misc & Non-ERP Expenses</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse border">
                        <thead>
                          <tr className="bg-gray-100">
                            <th className="border p-2 text-left">Type</th>
                            <th className="border p-2 text-left">Amount</th>
                            <th className="border p-2 text-left">Remarks</th>
                          </tr>
                        </thead>
                        <tbody>
                          {misc.map((item) => (
                            <tr key={item.id}>
                              <td className="border p-2 bg-gray-100">{item.type}</td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.amount}
                                  onChange={(e) =>
                                    setMisc((prev) =>
                                      prev.map((m) =>
                                        m.id === item.id ? { ...m, amount: Number.parseInt(e.target.value) } : m,
                                      ),
                                    )
                                  }
                                  className="w-32"
                                />
                              </td>
                              <td className="border p-2">
                                <Input
                                  value={item.remarks}
                                  onChange={(e) =>
                                    setMisc((prev) =>
                                      prev.map((m) => (m.id === item.id ? { ...m, remarks: e.target.value } : m)),
                                    )
                                  }
                                  className="w-40"
                                />
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>

              <TabsContent value="staff" className="space-y-6">
                <Card>
                  <CardHeader>
                    <CardTitle>Staff Salary</CardTitle>
                  </CardHeader>
                  <CardContent>
                    <div className="overflow-x-auto">
                      <table className="w-full border-collapse border">
                        <thead>
                          <tr className="bg-gray-100">
                            <th className="border p-2 text-left">Category</th>
                            <th className="border p-2 text-left">Nos</th>
                            <th className="border p-2 text-left">Salary per Month</th>
                            <th className="border p-2 text-left">No. of Months</th>
                            <th className="border p-2 text-left bg-gray-200">Salary Cost</th>
                          </tr>
                        </thead>
                        <tbody>
                          {staff.map((item) => (
                            <tr key={item.id}>
                              <td className="border p-2 bg-gray-100">{item.category}</td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.nos}
                                  onChange={(e) => {
                                    const newNos = Number.parseInt(e.target.value)
                                    setStaff((prev) =>
                                      prev.map((s) =>
                                        s.id === item.id
                                          ? { ...s, nos: newNos, salaryCost: newNos * s.salaryPerMonth * s.noOfMonths }
                                          : s,
                                      ),
                                    )
                                  }}
                                  className="w-20"
                                />
                              </td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.salaryPerMonth}
                                  onChange={(e) => {
                                    const newSalary = Number.parseInt(e.target.value)
                                    setStaff((prev) =>
                                      prev.map((s) =>
                                        s.id === item.id
                                          ? {
                                              ...s,
                                              salaryPerMonth: newSalary,
                                              salaryCost: s.nos * newSalary * s.noOfMonths,
                                            }
                                          : s,
                                      ),
                                    )
                                  }}
                                  className="w-24"
                                />
                              </td>
                              <td className="border p-2">
                                <Input
                                  type="number"
                                  value={item.noOfMonths}
                                  onChange={(e) => {
                                    const newMonths = Number.parseInt(e.target.value)
                                    setStaff((prev) =>
                                      prev.map((s) =>
                                        s.id === item.id
                                          ? {
                                              ...s,
                                              noOfMonths: newMonths,
                                              salaryCost: s.nos * s.salaryPerMonth * newMonths,
                                            }
                                          : s,
                                      ),
                                    )
                                  }}
                                  className="w-20"
                                  max="12"
                                />
                              </td>
                              <td className="border p-2 bg-gray-200">₹{item.salaryCost.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </CardContent>
                </Card>
              </TabsContent>
            </Tabs>

            {/* Save & Quit Button */}
            <div className="flex justify-center pt-6">
              <Button
                size="lg"
                className="bg-green-600 hover:bg-green-700"
                onClick={() =>
                  toast({
                    title: "Budget Saved",
                    description: "Budget will be prepared and written to Excel workbook. This may take a few minutes.",
                  })
                }
              >
                Save & Quit
              </Button>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  )
}
