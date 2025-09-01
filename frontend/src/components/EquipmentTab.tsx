import { useState, useEffect } from "react";
import {
  Button,
  Checkbox,
  Radio,
  Table,
  Space,
  message,
} from "antd";
import type { RadioChangeEvent } from "antd";
import type { ColumnsType } from "antd/es/table";
import { fetchData, postData, deleteData } from "../utils/api";

// --- INTERFACES ---
interface BPFixedExpense {
  id: number;
  Category: string;
  Cost: number;
}
interface FixedExpense {
  id: number;
  Category: string;
  Cost: number;
  remarks?: string;
}
interface HiredEquipment {
  id: number;
  Categoryname: string;
  EquipmentName: string;
  Make?: string;
  Model?: string;
  Capacity?: string;
  HireCharges: number;
  RAndMPercentage?: number;
  Hrs_PerMonth?: number;
  Fuel_PerHour?: number;
  Power_PerHour?: number;
  OperatorCost_PerMonth?: number;
}
interface LightingEquipment {
  id: number;
  Categoryname: string;
  EquipmentName: string;
  Capacity?: string;
  Make?: string;
  Model?: string;
  PowerPerUnit?: number;
  ConnectedLoad?: number;
  UtilityFactor?: number;
}
interface MajorEquipment {
  id: number;
  Categoryname: string;
  ConnectedLoadPerMC?: number;
  EquipmentName: string;
  Make?: string;
  Model?: string;
  Capacity?: string;
  Drive?: string;
  RepValue?: number;
  DepreciationPercentage?: number;
  Depreciation_Fixed?: number;
  Hrs_PerMonth?: number;
  Fuel_PerHour?: number;
  Power_PerHour?: number;
  OperatorCost_PerMonth?: number;
  RAndMPer_275?: number;
  RAndMPer_125?: number;
  RAndMPerc_050?: number;
  MaintCost_PerMonth?: number;
  PowerPerUnit_x0028_HP_x0029_?: number;
  UtilityFactor?: number;
}
interface MinorEquipment {
  id: number;
  Categoryname: string;
  EquipmentName: string;
  Make?: string;
  Model?: string;
  Capacity?: string;
  Drive?: string;
  CostOfNewEquipment?: number;
  RAndMPercentage?: number;
  Fuel_PerHour?: number;
  Power_PerHour?: number;
  DepreciationPercentage?: number;
  Hrs_PerMonth?: number;
  OperatorCost_PerMonth?: number;
  PowerPerUnit_x0028_HP_x0029_?: number;
  ConnectedLoadPerMC?: number;
  UtilityFactor?: number;
}

type AllEquipmentTypes =
  | MajorEquipment
  | MinorEquipment
  | HiredEquipment
  | LightingEquipment
  | FixedExpense
  | BPFixedExpense;

interface CombinedEquipmentItem {
  id: number;
  name?: string;
  unit?: string;
  selected?: boolean;
  mobDate?: string;
  demobDate?: string;
  qty?: number;
  shifts?: number;
  rate?: number;
  pVal?: number;
  hireCharges?: number;
  cost?: number;
  hrsMonth?: number;
  depreciation?: number;
  Category?: string;
  Cost?: number;
  remarks?: string;
  Categoryname?: string;
  EquipmentName?: string;
  Make?: string;
  Model?: string;
  Capacity?: string;
  HireCharges?: number;
  RAndMPercentage?: number;
  Hrs_PerMonth?: number;
  Fuel_PerHour?: number;
  Power_PerHour?: number;
  OperatorCost_PerMonth?: number;
  PowerPerUnit?: number;
  ConnectedLoad?: number;
  UtilityFactor?: number;
  ConnectedLoadPerMC?: number;
  Drive?: string;
  RepValue?: number;
  DepreciationPercentage?: number;
  Depreciation_Fixed?: number;
  RAndMPer_275?: number;
  RAndMPer_125?: number;
  RAndMPerc_050?: number;
  MaintCost_PerMonth?: number;
  PowerPerUnit_x0028_HP_x0029_?: number;
  CostOfNewEquipment?: number;
}

const EQUIPMENT_CATEGORY_MAP = {
  "Major Concrete": { endpoint: "majorequipments", type: "major" },
  "Major Conveyance": { endpoint: "majorequipments", type: "major" },
  "Major Crane": { endpoint: "majorequipments", type: "major" },
  "Major DG Sets": { endpoint: "majorequipments", type: "major" },
  "Major Material Handling": { endpoint: "majorequipments", type: "major" },
  "Major Non-Concrete": { endpoint: "majorequipments", type: "major" },
  "Major Others": { endpoint: "majorequipments", type: "major" },
  "Minor E Equipments": { endpoint: "minorequipments", type: "minor" },
  "Hired Equipments": { endpoint: "hiredequipments", type: "hired" },
  "Fixed Exp. - Tower Crane": { endpoint: "fixedexpenses", type: "fixed" },
  "Fixed Exp. - BP Related": { endpoint: "bpfixedexpenses", type: "bpFixed" },
  "Lighting/Single Phase Equips": {
    endpoint: "lightingequipments",
    type: "lighting",
  },
};

type EquipmentCategoryType = "major" | "minor" | "hired" | "lighting" | "fixed" | "bpFixed";

export default function EquipmentTabContent() {
  const [selectedCategory, setSelectedCategory] = useState("Major Concrete");
  const [equipment, setEquipment] = useState<CombinedEquipmentItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const mapBackendToUI = (data: AllEquipmentTypes[], type: EquipmentCategoryType): CombinedEquipmentItem[] => {
    let filteredData = data;
    console.log(data)
    if (type === "major" || type === "minor") {
      filteredData = (data as (MajorEquipment | MinorEquipment)[])
    }
    
    return filteredData.map((item) => {
        let nameField: string;
        let costField: string | null = null;
        let remarksField: string | null = null;

        switch (type) {
            case "major":
                nameField = "EquipmentName";
                costField = "RepValue";
                break;
            case "minor":
                nameField = "EquipmentName";
                costField = "CostOfNewEquipment";
                break;
            case "hired":
                nameField = "EquipmentName";
                costField = "HireCharges";
                break;
            case "lighting":
                nameField = "EquipmentName";
                break;
            case "fixed":
                nameField = "Category";
                costField = "Cost";
                remarksField = "remarks";
                break;
            case "bpFixed":
                nameField = "Category";
                costField = "Cost";
                break;
            default:
                nameField = "EquipmentName";
        }

        return {
            ...item,
            name: (item as any)[nameField],
            unit: "No",
            selected: false,
            mobDate: new Date().toISOString().slice(0, 10),
            demobDate: new Date().toISOString().slice(0, 10),
            qty: 1,
            shifts: 1,
            rate: costField ? (item as any)[costField] : 0,
            pVal: type === 'minor' && costField ? (item as any)[costField] : 0,
            hireCharges: type === 'hired' && costField ? (item as any)[costField] : 0,
            cost: (type === 'fixed' || type === 'bpFixed') && costField ? (item as any)[costField] : 0,
            hrsMonth: (item as any).Hrs_PerMonth || 100,
            depreciation: (item as any).DepreciationPercentage || 0,
            remarks: remarksField ? (item as any)[remarksField] : "",
        };
    });
  }

  useEffect(() => {
    console.log("useEffect triggered with:", selectedCategory);
    const loadEquipment = async () => {
      setLoading(true);
      setError(null);
      const categoryInfo = EQUIPMENT_CATEGORY_MAP[selectedCategory as keyof typeof EQUIPMENT_CATEGORY_MAP];
      console.log("Fetching data from endpoint:", categoryInfo?.endpoint);
      if (!categoryInfo) {
          setError("Category not found.");
          setLoading(false);
          return;
      }

      try {
        const data = await fetchData(categoryInfo.endpoint);
        const mappedData = mapBackendToUI(data, categoryInfo.type as EquipmentCategoryType);
        
        setEquipment(mappedData);
      } catch (err: any) {
        setError(err.message || `Failed to fetch data for ${selectedCategory}.`);
        message.error(err.message || `Failed to fetch data for ${selectedCategory}.`);
      } finally {
        setLoading(false);
      }
    };
    loadEquipment();
  }, [selectedCategory]);

//   const handleEquipmentUpdate = async (id: number, field: keyof CombinedEquipmentItem, value: any) => {
//     const categoryInfo = EQUIPMENT_CATEGORY_MAP[selectedCategory as keyof typeof EQUIPMENT_CATEGORY_MAP];
//     if (!categoryInfo) return;

//     const originalEquipment = [...equipment];
//     // Optimistically update UI state first
//     const updatedEquipment = equipment.map(item =>
//         item.id === id ? { ...item, [field]: value } : item
//     );
//     setEquipment(updatedEquipment);
    
//     // Only "rate" is persistent. Other fields like "qty" are session-only.
//     if (field !== 'rate') {
//         return;
//     }

//     const type = categoryInfo.type as EquipmentCategoryType;
//     let backendField: string | null = null;

//     switch (type) {
//         case 'major': backendField = 'RepValue'; break;
//         case 'minor': backendField = 'CostOfNewEquipment'; break;
//         case 'hired': backendField = 'HireCharges'; break;
//         case 'fixed': backendField = 'Cost'; break;
//         case 'bpFixed': backendField = 'Cost'; break;
//         case 'lighting': backendField = 'PowerPerUnit'; break;
//     }

//     if (!backendField) return;

//     try {
//         const payload = { [backendField]: value };
//         await putData(categoryInfo.endpoint, id, payload);
//         message.success("Rate updated successfully.");
//     } catch(err: any) {
//         message.error(`Update failed: ${err.message}. Reverting change.`);
//         // Revert optimistic update on failure
//         setEquipment(originalEquipment);
//     }
//   };

  const handleEquipmentToggle = (id: number) => {
    setEquipment(equipment.map(item => item.id === id ? {...item, selected: !item.selected} : item));
  };
  
  const handleSelectAll = (checked: boolean) => {
      setEquipment(equipment.map(item => ({...item, selected: checked})));
  };

  const getSelectedCount = () => equipment.filter(item => item.selected).length;

  // Simplified Add/Delete for example purposes
const addEquipment = async (baseItem: CombinedEquipmentItem) => {
  const categoryInfo = EQUIPMENT_CATEGORY_MAP[selectedCategory as keyof typeof EQUIPMENT_CATEGORY_MAP];
  if (!categoryInfo) return;

  // Prepare payload
  const newItemPayload = { ...baseItem, EquipmentName: `${baseItem.EquipmentName} (Copy)` };
  delete (newItemPayload as any).id;

  try {
    // Call backend
    const addedItem = await postData(categoryInfo.endpoint, newItemPayload);

    // Build new row
    const newRow = { ...addedItem, selected: false, name: addedItem.EquipmentName };

    // Find index of clicked row
    const index = equipment.findIndex((item) => item.id === baseItem.id);

    if (index === -1) {
      // fallback: just append if not found
      setEquipment([...equipment, newRow]);
    } else {
      // insert right below clicked row
      const updated = [...equipment];
      updated.splice(index + 1, 0, newRow);
      setEquipment(updated);
    }

    message.success("Equipment duplicated right below.");
  } catch (err: any) {
    message.error(`Failed to duplicate: ${err.message}`);
  }
};



  const deleteEquipment = async (id: number) => {
    const categoryInfo = EQUIPMENT_CATEGORY_MAP[selectedCategory as keyof typeof EQUIPMENT_CATEGORY_MAP];
    if (!categoryInfo) return;
    try {
        await deleteData(categoryInfo.endpoint, id);
        setEquipment(equipment.filter(item => item.id !== id));
        message.success("Equipment deleted.");
    } catch (err: any) {
        message.error(`Failed to delete: ${err.message}`);
    }
  };
const actionColumn: ColumnsType<any>[number] = {
  title: "Action",
  key: "action",
  width: 150,
  fixed: "right" as const,   // ✅ ensures it's 'right' not string
  render: (_: any, record: any) => (
    <Space size="middle">
      <Button onClick={() => addEquipment(record)}>Add</Button>
      <Button danger onClick={() => deleteEquipment(record.id)}>Delete</Button>
    </Space>
  ),
};


const selectColumn = {
  title: "Select",
  dataIndex: "selected",
  key: "selected",
  width: 70,
  render: (_: any, record: any) => (
    <Checkbox
      checked={record.selected}
      onChange={() => handleEquipmentToggle(record.id)}
    />
  ),
};

const getColumnsByType = (type: EquipmentCategoryType): ColumnsType<any> => {
  switch (type) {
    case "bpFixed":
      return [
        selectColumn,
        { title: "Category", dataIndex: "Category", key: "Category" },
        { title: "Cost", dataIndex: "Cost", key: "Cost" },
        actionColumn,
      ];

    case "fixed":
      return [
        selectColumn,
        { title: "Category", dataIndex: "Category", key: "Category" },
        { title: "Cost", dataIndex: "Cost", key: "Cost" },
        { title: "Remarks", dataIndex: "remarks", key: "remarks" },
        actionColumn,
      ];

    case "hired":
      return [
        selectColumn,
        { title: "Category", dataIndex: "Categoryname", key: "Categoryname" },
        { title: "Equipment Name", dataIndex: "EquipmentName", key: "EquipmentName" },
        { title: "Make", dataIndex: "Make", key: "Make" },
        { title: "Model", dataIndex: "Model", key: "Model" },
        { title: "Capacity", dataIndex: "Capacity", key: "Capacity" },
        { title: "Hire Charges", dataIndex: "HireCharges", key: "HireCharges" },
        { title: "R&M %", dataIndex: "RAndMPercentage", key: "RAndMPercentage" },
        { title: "Hrs/Month", dataIndex: "Hrs_PerMonth", key: "Hrs_PerMonth" },
        { title: "Fuel/Hour", dataIndex: "Fuel_PerHour", key: "Fuel_PerHour" },
        { title: "Power/Hour", dataIndex: "Power_PerHour", key: "Power_PerHour" },
        { title: "Operator Cost/Month", dataIndex: "OperatorCost_PerMonth", key: "OperatorCost_PerMonth" },
        actionColumn,
      ];

    case "lighting":
      return [
        selectColumn,
        { title: "Category", dataIndex: "Categoryname", key: "Categoryname" },
        { title: "Equipment Name", dataIndex: "EquipmentName", key: "EquipmentName" },
        { title: "Capacity", dataIndex: "Capacity", key: "Capacity" },
        { title: "Make", dataIndex: "Make", key: "Make" },
        { title: "Model", dataIndex: "Model", key: "Model" },
        { title: "Power Per Unit", dataIndex: "PowerPerUnit", key: "PowerPerUnit" },
        { title: "Connected Load", dataIndex: "ConnectedLoad", key: "ConnectedLoad" },
        { title: "Utility Factor", dataIndex: "UtilityFactor", key: "UtilityFactor" },
        actionColumn,
      ];

    case "major":
      return [
        selectColumn,
        { title: "Category", dataIndex: "Categoryname", key: "Categoryname" },
        { title: "Equipment Name", dataIndex: "EquipmentName", key: "EquipmentName" },
        { title: "Make", dataIndex: "Make", key: "Make" },
        { title: "Model", dataIndex: "Model", key: "Model" },
        { title: "Capacity", dataIndex: "Capacity", key: "Capacity" },
        { title: "Drive", dataIndex: "Drive", key: "Drive" },
        { title: "Rep Value", dataIndex: "RepValue", key: "RepValue" },
        { title: "Depreciation %", dataIndex: "DepreciationPercentage", key: "DepreciationPercentage" },
        { title: "Depreciation Fixed", dataIndex: "Depreciation_Fixed", key: "Depreciation_Fixed" },
        { title: "Hrs/Month", dataIndex: "Hrs_PerMonth", key: "Hrs_PerMonth" },
        { title: "Fuel/Hour", dataIndex: "Fuel_PerHour", key: "Fuel_PerHour" },
        { title: "Power/Hour", dataIndex: "Power_PerHour", key: "Power_PerHour" },
        { title: "Operator Cost/Month", dataIndex: "OperatorCost_PerMonth", key: "OperatorCost_PerMonth" },
        { title: "R&M 275%", dataIndex: "RAndMPer_275", key: "RAndMPer_275" },
        { title: "R&M 125%", dataIndex: "RAndMPer_125", key: "RAndMPer_125" },
        { title: "R&M 50%", dataIndex: "RAndMPerc_050", key: "RAndMPerc_050" },
        { title: "Maint Cost/Month", dataIndex: "MaintCost_PerMonth", key: "MaintCost_PerMonth" },
        { title: "Power (HP)", dataIndex: "PowerPerUnit_x0028_HP_x0029_", key: "PowerPerUnit_x0028_HP_x0029_" },
        { title: "Utility Factor", dataIndex: "UtilityFactor", key: "UtilityFactor" },
        { title: "Connected Load/MC", dataIndex: "ConnectedLoadPerMC", key: "ConnectedLoadPerMC" },
        actionColumn,
      ];

    case "minor":
      return [
        selectColumn,
        { title: "Category", dataIndex: "Categoryname", key: "Categoryname" },
        { title: "Equipment Name", dataIndex: "EquipmentName", key: "EquipmentName" },
        { title: "Make", dataIndex: "Make", key: "Make" },
        { title: "Model", dataIndex: "Model", key: "Model" },
        { title: "Capacity", dataIndex: "Capacity", key: "Capacity" },
        { title: "Drive", dataIndex: "Drive", key: "Drive" },
        { title: "Cost of New Equipment", dataIndex: "CostOfNewEquipment", key: "CostOfNewEquipment" },
        { title: "R&M %", dataIndex: "RAndMPercentage", key: "RAndMPercentage" },
        { title: "Fuel/Hour", dataIndex: "Fuel_PerHour", key: "Fuel_PerHour" },
        { title: "Power/Hour", dataIndex: "Power_PerHour", key: "Power_PerHour" },
        { title: "Depreciation %", dataIndex: "DepreciationPercentage", key: "DepreciationPercentage" },
        { title: "Hrs/Month", dataIndex: "Hrs_PerMonth", key: "Hrs_PerMonth" },
        { title: "Operator Cost/Month", dataIndex: "OperatorCost_PerMonth", key: "OperatorCost_PerMonth" },
        { title: "Power (HP)", dataIndex: "PowerPerUnit_x0028_HP_x0029_", key: "PowerPerUnit_x0028_HP_x0029_" },
        { title: "Connected Load/MC", dataIndex: "ConnectedLoadPerMC", key: "ConnectedLoadPerMC" },
        { title: "Utility Factor", dataIndex: "UtilityFactor", key: "UtilityFactor" },
        actionColumn,
      ];

    default:
      return [selectColumn, { title: "Name", dataIndex: "name", key: "name" }, actionColumn];
  }
};

// const equipmentColumns = (): ColumnsType<CombinedEquipmentItem> => {
//   const categoryInfo = EQUIPMENT_CATEGORY_MAP[selectedCategory as keyof typeof EQUIPMENT_CATEGORY_MAP];
//   if (!categoryInfo) return [];

//   const type = categoryInfo.type as EquipmentCategoryType;

//   const baseColumns: ColumnsType<CombinedEquipmentItem> = [
//     {
//       title: "Select",
//       dataIndex: "selected",
//       key: "selected",
//       width: 70,
//       render: (_, record) => (
//         <Checkbox checked={record.selected} onChange={() => handleEquipmentToggle(record.id)} />
//       ),
//     },
//     { title: "Name", dataIndex: "name", key: "name", width: 250, fixed: "left" },
//   ];

//   let extraColumns: ColumnsType<CombinedEquipmentItem> = [];

//   switch (type) {
//     case "major":
//     case "minor":
//       extraColumns = [
//         { title: "Capacity", dataIndex: "Capacity", key: "Capacity" },
//         { title: "Make", dataIndex: "Make", key: "Make" },
//         { title: "Model", dataIndex: "Model", key: "Model" },
//         { title: "Rep Value", dataIndex: "RepValue", key: "RepValue" },
//         { title: "Depreciation %", dataIndex: "DepreciationPercentage", key: "DepreciationPercentage" },
//       ];
//       break;

//     case "hired":
//       extraColumns = [
//         { title: "Hire Charges", dataIndex: "HireCharges", key: "HireCharges" },
//         { title: "Hrs/Month", dataIndex: "Hrs_PerMonth", key: "Hrs_PerMonth" },
//       ];
//       break;

//     case "lighting":
//       extraColumns = [
//         { title: "Power Per Unit", dataIndex: "PowerPerUnit", key: "PowerPerUnit" },
//         { title: "Utility Factor", dataIndex: "UtilityFactor", key: "UtilityFactor" },
//       ];
//       break;

//     case "fixed":
//     case "bpFixed":
//       extraColumns = [
//         { title: "Category", dataIndex: "Category", key: "Category" },
//         { title: "Cost", dataIndex: "Cost", key: "Cost" },
//         ...(type === "fixed" ? [{ title: "Remarks", dataIndex: "remarks", key: "remarks" }] : []),
//       ];
//       break;
//   }

//   const actionColumn: ColumnsType<CombinedEquipmentItem>[number] = {
//     title: "Action",
//     key: "action",
//     width: 150,
//     fixed: "right",
//     render: (_, record) => (
//       <Space size="middle">
//         <Button onClick={() => addEquipment(record)}>Add</Button>
//         <Button danger onClick={() => deleteEquipment(record.id)}>Delete</Button>
//       </Space>
//     ),
//   };

//   return [...baseColumns, ...extraColumns, actionColumn];
// };
const categoryInfo = EQUIPMENT_CATEGORY_MAP[selectedCategory as keyof typeof EQUIPMENT_CATEGORY_MAP];


  return (
    <div className="space-y-6">
      <div className="flex gap-4 justify-center flex-wrap">
        <Button onClick={() => handleSelectAll(true)}>Select All</Button>
        <Button onClick={() => handleSelectAll(false)}>Unselect All</Button>
        <Button>View Selected Items List ({getSelectedCount()})</Button>
      </div>

      <div className="border rounded-lg overflow-hidden">
        <div className="bg-blue-100 p-4">
          <h3 className="font-semibold text-lg">{selectedCategory} Equipment</h3>
        </div>
        {loading ? (
          <div className="p-4 text-center">Loading equipment data...</div>
        ) : error ? (
          <div className="p-4 text-center text-red-500">{error}</div>
        ) : (
            <Table
            dataSource={equipment}
            columns={getColumnsByType(categoryInfo.type as EquipmentCategoryType)}
            rowKey="id"
            pagination={false}
            bordered
            scroll={{ x: "max-content" }}
            />

        )}
      </div>

      <div className="border-t pt-6">
        <h4 className="font-semibold mb-4">Equipment Categories:</h4>
        <Radio.Group
          onChange={(e: RadioChangeEvent) => setSelectedCategory(e.target.value)}
          value={selectedCategory}
        >
          <Space direction="horizontal" wrap>
            {Object.keys(EQUIPMENT_CATEGORY_MAP).map((category) => (
              <Radio key={category} value={category}>
                {category}
              </Radio>
            ))}
          </Space>
        </Radio.Group>
      </div>
    </div>
  );
}