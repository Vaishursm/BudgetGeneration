"use client";

import { useState } from "react";
import {
  Button,
  Tabs,
  Input,
  Card,
  Table,
  ConfigProvider,
  message,
} from "antd";
import type { TabsProps } from "antd";
import type { ColumnsType } from "antd/es/table";
import EquipmentTab from "./EquipmentTab"; // Import the new component

interface ElectricalData {
  installationPercent: number;
  breakup: {
    cabling: number;
    switchgear: number;
    lighting: number;
    others: number;
  };
}

const electricalData: ElectricalData = {
  installationPercent: 15,
  breakup: {
    cabling: 40,
    switchgear: 25,
    lighting: 20,
    others: 15,
  },
};

interface PipelineExpense {
  id: number;
  category: string;
  qty: number;
  costPerUnit: number;
  amount: number;
  remarks: string;
}

const pipelineExpenses: PipelineExpense[] = [
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
];

interface ElectMechanicCost {
  id: number;
  category: string;
  nos: number;
  salaryPerMonth: number;
  noOfMonths: number;
  salaryCost: number;
}

const electMechanicCost: ElectMechanicCost[] = [
  { id: 1, category: "Electrician", nos: 3, salaryPerMonth: 25000, noOfMonths: 12, salaryCost: 900000 },
  { id: 2, category: "Mechanic", nos: 2, salaryPerMonth: 22000, noOfMonths: 12, salaryCost: 528000 },
  { id: 3, category: "Welder", nos: 4, salaryPerMonth: 20000, noOfMonths: 10, salaryCost: 800000 },
  { id: 4, category: "Helper", nos: 6, salaryPerMonth: 15000, noOfMonths: 12, salaryCost: 1080000 },
];

interface MiscExpense {
  id: number;
  type: string;
  amount: number;
  remarks: string;
}

const miscExpenses: MiscExpense[] = [
  { id: 1, type: "Insurance", amount: 150000, remarks: "Equipment insurance" },
  { id: 2, type: "Transportation", amount: 200000, remarks: "Equipment transportation" },
  { id: 3, type: "Permits & Licenses", amount: 75000, remarks: "Various permits" },
  { id: 4, type: "Safety Equipment", amount: 100000, remarks: "Safety gear and equipment" },
];

interface StaffSalary {
  id: number;
  category: string;
  nos: number;
  salaryPerMonth: number;
  noOfMonths: number;
  salaryCost: number;
}

const staffSalary: StaffSalary[] = [
  { id: 1, category: "Project Manager", nos: 1, salaryPerMonth: 80000, noOfMonths: 12, salaryCost: 960000 },
  { id: 2, category: "Site Engineer", nos: 2, salaryPerMonth: 45000, noOfMonths: 12, salaryCost: 1080000 },
  { id: 3, category: "Supervisor", nos: 3, salaryPerMonth: 30000, noOfMonths: 12, salaryCost: 1080000 },
  { id: 4, category: "Safety Officer", nos: 1, salaryPerMonth: 35000, noOfMonths: 12, salaryCost: 420000 },
];

export default function MainInterface() {
  const [activeTab, setActiveTab] = useState("equipment");
  const [electrical, setElectrical] = useState(electricalData);
  const [pipeline, setPipeline] = useState(pipelineExpenses);
  const [electMechanic, setElectMechanic] = useState(electMechanicCost);
  const [misc, setMisc] = useState(miscExpenses);
  const [staff, setStaff] = useState(staffSalary);

  const pipelineColumns: ColumnsType<PipelineExpense> = [
    {
      title: "Category",
      dataIndex: "category",
      key: "category",
      width: 150,
      render: (text) => <span className="font-medium">{text}</span>,
    },
    {
      title: "Qty",
      dataIndex: "qty",
      key: "qty",
      width: 100,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newQty = Number.parseInt(e.target.value);
            setPipeline((prev) =>
              prev.map((p) => (p.id === record.id ? { ...p, qty: newQty, amount: newQty * p.costPerUnit } : p)),
            );
          }}
          min={1}
        />
      ),
    },
    {
      title: "Cost per Unit",
      dataIndex: "costPerUnit",
      key: "costPerUnit",
      width: 150,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newCost = Number.parseInt(e.target.value);
            setPipeline((prev) =>
              prev.map((p) => (p.id === record.id ? { ...p, costPerUnit: newCost, amount: p.qty * newCost } : p)),
            );
          }}
          min={0}
        />
      ),
    },
    {
      title: "Amount",
      dataIndex: "amount",
      key: "amount",
      width: 150,
      render: (text) => `₹${text.toLocaleString()}`,
    },
    {
      title: "Remarks",
      dataIndex: "remarks",
      key: "remarks",
      width: 200,
      render: (text, record) => (
        <Input
          value={text}
          onChange={(e) =>
            setPipeline((prev) => prev.map((p) => (p.id === record.id ? { ...p, remarks: e.target.value } : p)))
          }
        />
      ),
    },
  ];

  const electMechanicColumns: ColumnsType<ElectMechanicCost> = [
    {
      title: "Category",
      dataIndex: "category",
      key: "category",
      width: 150,
      render: (text) => <span className="font-medium">{text}</span>,
    },
    {
      title: "Nos",
      dataIndex: "nos",
      key: "nos",
      width: 100,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newNos = Number.parseInt(e.target.value);
            setElectMechanic((prev) =>
              prev.map((em) =>
                em.id === record.id
                  ? {
                      ...em,
                      nos: newNos,
                      salaryCost: newNos * em.salaryPerMonth * em.noOfMonths,
                    }
                  : em,
              ),
            );
          }}
          min={0}
        />
      ),
    },
    {
      title: "Salary per Month",
      dataIndex: "salaryPerMonth",
      key: "salaryPerMonth",
      width: 150,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newSalary = Number.parseInt(e.target.value);
            setElectMechanic((prev) =>
              prev.map((em) =>
                em.id === record.id
                  ? {
                      ...em,
                      salaryPerMonth: newSalary,
                      salaryCost: em.nos * newSalary * em.noOfMonths,
                    }
                  : em,
              ),
            );
          }}
          min={0}
        />
      ),
    },
    {
      title: "No. of Months",
      dataIndex: "noOfMonths",
      key: "noOfMonths",
      width: 150,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newMonths = Number.parseInt(e.target.value);
            setElectMechanic((prev) =>
              prev.map((em) =>
                em.id === record.id
                  ? {
                      ...em,
                      noOfMonths: newMonths,
                      salaryCost: em.nos * em.salaryPerMonth * newMonths,
                    }
                  : em,
              ),
            );
          }}
          min={0}
          max={12}
        />
      ),
    },
    {
      title: "Salary Cost",
      dataIndex: "salaryCost",
      key: "salaryCost",
      width: 150,
      render: (text) => `₹${text.toLocaleString()}`,
    },
  ];

  const miscColumns: ColumnsType<MiscExpense> = [
    {
      title: "Type",
      dataIndex: "type",
      key: "type",
      width: 200,
      render: (text) => <span className="font-medium">{text}</span>,
    },
    {
      title: "Amount",
      dataIndex: "amount",
      key: "amount",
      width: 150,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) =>
            setMisc((prev) => prev.map((m) => (m.id === record.id ? { ...m, amount: Number.parseInt(e.target.value) } : m)))
          }
          min={0}
        />
      ),
    },
    {
      title: "Remarks",
      dataIndex: "remarks",
      key: "remarks",
      width: 250,
      render: (text, record) => (
        <Input
          value={text}
          onChange={(e) =>
            setMisc((prev) => prev.map((m) => (m.id === record.id ? { ...m, remarks: e.target.value } : m)))
          }
        />
      ),
    },
  ];

  const staffColumns: ColumnsType<StaffSalary> = [
    {
      title: "Category",
      dataIndex: "category",
      key: "category",
      width: 150,
      render: (text) => <span className="font-medium">{text}</span>,
    },
    {
      title: "Nos",
      dataIndex: "nos",
      key: "nos",
      width: 100,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newNos = Number.parseInt(e.target.value);
            setStaff((prev) =>
              prev.map((s) =>
                s.id === record.id
                  ? { ...s, nos: newNos, salaryCost: newNos * s.salaryPerMonth * s.noOfMonths }
                  : s,
              ),
            );
          }}
          min={0}
        />
      ),
    },
    {
      title: "Salary per Month",
      dataIndex: "salaryPerMonth",
      key: "salaryPerMonth",
      width: 150,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newSalary = Number.parseInt(e.target.value);
            setStaff((prev) =>
              prev.map((s) =>
                s.id === record.id
                  ? {
                      ...s,
                      salaryPerMonth: newSalary,
                      salaryCost: s.nos * newSalary * s.noOfMonths,
                    }
                  : s,
              ),
            );
          }}
          min={0}
        />
      ),
    },
    {
      title: "No. of Months",
      dataIndex: "noOfMonths",
      key: "noOfMonths",
      width: 150,
      render: (text, record) => (
        <Input
          type="number"
          value={text}
          onChange={(e) => {
            const newMonths = Number.parseInt(e.target.value);
            setStaff((prev) =>
              prev.map((s) =>
                s.id === record.id
                  ? {
                      ...s,
                      noOfMonths: newMonths,
                      salaryCost: s.nos * s.salaryPerMonth * newMonths,
                    }
                  : s,
              ),
            );
          }}
          min={0}
          max={12}
        />
      ),
    },
    {
      title: "Salary Cost",
      dataIndex: "salaryCost",
      key: "salaryCost",
      width: 150,
      render: (text) => `₹${text.toLocaleString()}`,
    },
  ];

  const items: TabsProps["items"] = [
    {
      key: "equipment",
      label: "Equipment Entry",
      children: <EquipmentTab />,
    },
    {
      key: "electricals",
      label: "Electricals",
      children: (
        <Card>
          <Card.Meta title={<span className="text-xl font-bold">Electrical Installation</span>} />
          <div className="space-y-4 mt-4">
            <div>
              <label className="block text-sm font-medium text-gray-700">Electrical Installation % at the top</label>
              <Input
                type="number"
                value={electrical.installationPercent}
                onChange={(e) =>
                  setElectrical((prev) => ({ ...prev, installationPercent: Number.parseFloat(e.target.value) }))
                }
                className="w-32 mt-1"
                min={0}
                max={100}
                step={0.1}
              />
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700">Cabling %</label>
                <Input
                  type="number"
                  value={electrical.breakup.cabling}
                  onChange={(e) =>
                    setElectrical((prev) => ({
                      ...prev,
                      breakup: { ...prev.breakup, cabling: Number.parseFloat(e.target.value) },
                    }))
                  }
                  className="w-24 mt-1"
                  min={0}
                  max={100}
                  step={0.1}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Switchgear %</label>
                <Input
                  type="number"
                  value={electrical.breakup.switchgear}
                  onChange={(e) =>
                    setElectrical((prev) => ({
                      ...prev,
                      breakup: { ...prev.breakup, switchgear: Number.parseFloat(e.target.value) },
                    }))
                  }
                  className="w-24 mt-1"
                  min={0}
                  max={100}
                  step={0.1}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Lighting %</label>
                <Input
                  type="number"
                  value={electrical.breakup.lighting}
                  onChange={(e) =>
                    setElectrical((prev) => ({
                      ...prev,
                      breakup: { ...prev.breakup, lighting: Number.parseFloat(e.target.value) },
                    }))
                  }
                  className="w-24 mt-1"
                  min={0}
                  max={100}
                  step={0.1}
                />
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700">Others %</label>
                <Input
                  type="number"
                  value={electrical.breakup.others}
                  onChange={(e) =>
                    setElectrical((prev) => ({
                      ...prev,
                      breakup: { ...prev.breakup, others: Number.parseFloat(e.target.value) },
                    }))
                  }
                  className="w-24 mt-1"
                  min={0}
                  max={100}
                  step={0.1}
                />
              </div>
            </div>
            <div className="flex gap-4 mt-6">
              <Button onClick={() => setActiveTab("equipment")}>Goto Equipment Entry</Button>
              <Button onClick={() => setActiveTab("pipeline")}>Goto Pipeline Expenses Entry</Button>
            </div>
          </div>
        </Card>
      ),
    },
    {
      key: "pipeline",
      label: "Pipeline Expenses",
      children: (
        <Card>
          <Card.Meta title={<span className="text-xl font-bold">Pipeline Expenses</span>} />
          <div className="mt-4">
            <Table
              dataSource={pipeline}
              columns={pipelineColumns}
              rowKey="id"
              pagination={false}
              bordered
              scroll={{ x: "max-content" }}
            />
          </div>
        </Card>
      ),
    },
    {
      key: "electmechanic",
      label: "Elect/Mechanic Cost",
      children: (
        <Card>
          <Card.Meta title={<span className="text-xl font-bold">Elect/Mechanic Cost</span>} />
          <div className="mt-4">
            <Table
              dataSource={electMechanic}
              columns={electMechanicColumns}
              rowKey="id"
              pagination={false}
              bordered
              scroll={{ x: "max-content" }}
            />
          </div>
        </Card>
      ),
    },
    {
      key: "misc",
      label: "Misc & Non-ERP",
      children: (
        <Card>
          <Card.Meta title={<span className="text-xl font-bold">Misc & Non-ERP Expenses</span>} />
          <div className="mt-4">
            <Table
              dataSource={misc}
              columns={miscColumns}
              rowKey="id"
              pagination={false}
              bordered
              scroll={{ x: "max-content" }}
            />
          </div>
        </Card>
      ),
    },
    {
      key: "staff",
      label: "Staff Salary",
      children: (
        <Card>
          <Card.Meta title={<span className="text-xl font-bold">Staff Salary</span>} />
          <div className="mt-4">
            <Table
              dataSource={staff}
              columns={staffColumns}
              rowKey="id"
              pagination={false}
              bordered
              scroll={{ x: "max-content" }}
            />
          </div>
        </Card>
      ),
    },
  ];

  return (
    <ConfigProvider
      theme={{
        token: {
          colorPrimary: "#1890ff", // Ant Design's default primary blue
        },
      }}
    >
      <div className="min-h-screen bg-gray-50 p-4">
        <div className="max-w-full mx-auto">
          <Card
            title={<div className="text-2xl font-bold text-center">Budget Generation - Main Interface</div>}
            className="shadow-lg"
          >
            <Tabs
              activeKey={activeTab}
              onChange={setActiveTab}
              items={items}
              tabBarStyle={{ justifyContent: "center" }}
            />

            {/* Save & Quit Button */}
            <div className="flex justify-center pt-6">
              <Button
                size="large"
                type="primary"
                className="bg-green-600 hover:bg-green-700"
                onClick={() =>
                  message.success(
                    "Budget will be prepared and written to Excel workbook. This may take a few minutes."
                  )
                }
              >
                Save & Quit
              </Button>
            </div>
          </Card>
        </div>
      </div>
    </ConfigProvider>
  );
}
