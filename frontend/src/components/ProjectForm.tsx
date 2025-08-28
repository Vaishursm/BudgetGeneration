import { useState, useEffect } from "react";
import {
  Form,
  Input,
  InputNumber,
  DatePicker,
  Button,
  Select,
  message,
  notification,
} from "antd";
import { UploadOutlined } from "@ant-design/icons";
import { useForm, Controller, type SubmitHandler } from "react-hook-form";
import * as yup from "yup";
import { yupResolver } from "@hookform/resolvers/yup";
import dayjs, { Dayjs } from "dayjs";
import axios from "axios";

const { TextArea } = Input;
const { Option } = Select;

interface ProjectFormValues {
  chooseProject?: string;
  projectCode: string;
  description: string;
  clientName: string;
  projectLocation: string;
  projectValue: number | null;
  startDate: Dayjs | null;
  endDate: Dayjs | null;
  concreteQty: number | null;
  fuelCost: number | null;
  powerCost: number | null;
  filePath: string;
}

const schema: yup.ObjectSchema<ProjectFormValues> = yup.object().shape({
  chooseProject: yup.string().optional(),
  projectCode: yup.string().required("Project code is required"),
  description: yup.string().required("Description is required"),
  clientName: yup.string().required("Client name is required"),
  projectLocation: yup.string().required("Location is required"),
  projectValue: yup
    .number()
    .typeError("Must be a number")
    .positive("Must be positive")
    .required("Project value is required"),
  startDate: yup
    .mixed<Dayjs>()
    .required("Start date is required")
    .nullable()
    .test("is-future", "Start date must be in the future", (value) =>
      value ? dayjs(value).isAfter(dayjs(), "day") : false
    ),
  endDate: yup
    .mixed<Dayjs>()
    .required("End date is required")
    .nullable()
    .test("is-after", "End date must be after start date", function (value) {
      const { startDate } = this.parent as ProjectFormValues;
      return startDate && value ? dayjs(value).isAfter(startDate) : true;
    }),
  concreteQty: yup
    .number()
    .typeError("Must be a number")
    .integer("Decimals not allowed")
    .min(0, "Must be >= 0")
    .required("Quantity is required"),
  fuelCost: yup
    .number()
    .typeError("Must be a number")
    .positive("Must be positive")
    .required("Fuel cost is required"),
  powerCost: yup
    .number()
    .typeError("Must be a number")
    .positive("Must be positive")
    .required("Power cost is required"),
  filePath: yup
    .string()
    .required("Workbook path is required")
    .matches(/\.(xlsx|xls)$/, "Must be .xlsx or .xls file"),
});

export default function ProjectForm() {
  const [isNewProject, setIsNewProject] = useState(true);
  const [existingProjects, setExistingProjects] = useState<
    Record<string, ProjectFormValues>
  >({});
  const [api, contextHolder] = notification.useNotification();

  const openNotification = (
    placement: "topLeft" | "topRight" | "bottomLeft" | "bottomRight" | "top" | "bottom",
    description: string,
    message: string,
    type: "success" | "error" | "warning" | "info"
  ) => {
    api[type]({
      message,
      description,
      placement,
      duration: 3,
      showProgress :true,
      pauseOnHover: true,
      closeIcon: true,
    });
  };

  const {
    handleSubmit,
    control,
    reset,
    formState: { errors },
  } = useForm<ProjectFormValues>({
    resolver: yupResolver(schema),
    defaultValues: {
      chooseProject: "new",
      projectCode: "",
      description: "",
      clientName: "",
      projectLocation: "",
      projectValue: null,
      startDate: null,
      endDate: null,
      concreteQty: null,
      fuelCost: null,
      powerCost: null,
      filePath: "",
    },
  });

  interface ProjectFromAPI {
    id: number;
    projectCode: string;
    description: string;
    clientName: string;
    projectLocation: string;
    projectValue: number;
    startDate: string | null;
    endDate: string | null;
    concreteQty: number;
    fuelCost: number;
    powerCost: number;
    filePath: string;
  }

  // ✅ Fetch existing projects from API
  useEffect(() => {
    axios
      .get("http://localhost:5000/projects")
      .then((res) => {
        const projectsData: Record<string, ProjectFormValues> = {};
        res.data.forEach((proj: ProjectFromAPI) => {
          projectsData[`project${proj.id}`] = {
            chooseProject: `project${proj.id}`,
            projectCode: proj.projectCode,
            description: proj.description,
            clientName: proj.clientName,
            projectLocation: proj.projectLocation,
            projectValue: proj.projectValue,
            startDate: proj.startDate ? dayjs(proj.startDate) : null,
            endDate: proj.endDate ? dayjs(proj.endDate) : null,
            concreteQty: proj.concreteQty,
            fuelCost: proj.fuelCost,
            powerCost: proj.powerCost,
            filePath: proj.filePath,
          };
        });
        setExistingProjects(projectsData);
        setIsNewProject(true);
      })
      .catch((err: unknown) => {
        console.error(err);
        openNotification("top", "Failed to fetch projects", "Error", "error");
      });
    //eslint-disable-next-line
  }, []);

  const onSubmit: SubmitHandler<ProjectFormValues> = async (data) => {
    try {
      if (isNewProject) {
        // POST request to create a new project
        await axios.post("http://localhost:5000/projects", {
          ...data,
          startDate: data.startDate?.toISOString(),
          endDate: data.endDate?.toISOString(),
        });
        openNotification("top", "Project updated successfully!", "Success", "success");
        reset({ chooseProject: "new" });
        setIsNewProject(true);  
      } else {
        // PUT request to update existing project
        console.log(data)
        const projectId = data.chooseProject?.replace("project", "");
        await axios.put(`http://localhost:5000/projects/${projectId}`, {
          ...data,
          startDate: data.startDate?.toISOString(),
          endDate: data.endDate?.toISOString(),
        });
        openNotification("top", "Project updated successfully!", "Success", "success");
        reset({ chooseProject: "new" });
        setIsNewProject(true);
      }
      // Refresh projects list
      const res = await axios.get("http://localhost:5000/projects");
      const projectsData: Record<string, ProjectFormValues> = {};
      res.data.forEach((proj: ProjectFromAPI) => {
        projectsData[`project${proj.id}`] = {
          chooseProject: `project${proj.id}`,
          projectCode: proj.projectCode,
          description: proj.description,
          clientName: proj.clientName,
          projectLocation: proj.projectLocation,
          projectValue: proj.projectValue,
          startDate: proj.startDate ? dayjs(proj.startDate) : null,
          endDate: proj.endDate ? dayjs(proj.endDate) : null,
          concreteQty: proj.concreteQty,
          fuelCost: proj.fuelCost,
          powerCost: proj.powerCost,
          filePath: proj.filePath,
        };
      });
      setExistingProjects(projectsData);
      setIsNewProject(true);
    } catch (err: unknown) {
        if (axios.isAxiosError(err)) {
          // ✅ err is AxiosError here
          openNotification(
            "top",
            err.response?.data?.error || err.message || "Failed to save project",
            "Error",
            "error"
          );
        } else if (err instanceof Error) {
          // ✅ plain error
          openNotification("top", err.message, "Error", "error");
        } else {
          openNotification("top", "An unexpected error occurred", "Error", "error");
        }
    }
  };

  return (
    <div className="project-card">
      {contextHolder}
      <Form layout="vertical" onFinish={handleSubmit(onSubmit)}>
        {/* Choose Project */}
        <Form.Item label="Which Project you want to open to work with">
          <Controller
            name="chooseProject"
            control={control}
            render={({ field }) => (
              <Select
                {...field}
                onChange={(val) => {
                  field.onChange(val);
                  if (val === "new") {
                    setIsNewProject(true);
                    reset({ chooseProject: "new" });
                  } else {
                    setIsNewProject(false);
                    if (existingProjects[val]) {
                      reset(existingProjects[val]);
                    }
                  }
                }}
              >
                <Option value="new">+ Create Project</Option>
                {Object.keys(existingProjects).map((key) => (
                  <Option key={key} value={key}>
                    {existingProjects[key].projectCode} -{" "}
                    {existingProjects[key].description}
                  </Option>
                ))}
              </Select>
            )}
          />
        </Form.Item>
        {/* Project Code */}
        <Form.Item
          label="Project Code"
          className="project-code"
          validateStatus={errors.projectCode ? "error" : ""}
          help={errors.projectCode?.message}
        >
          <Controller
            name="projectCode"
            control={control}
            render={({ field }) => (
              <Input
                {...field}
                disabled={!isNewProject}
                placeholder="Enter project code"
              />
            )}
          />
        </Form.Item>

        {/* Description */}
        <Form.Item
          label="Description"
          validateStatus={errors.description ? "error" : ""}
          help={errors.description?.message}
        >
          <Controller
            name="description"
            control={control}
            render={({ field }) => (
              <TextArea {...field} disabled={!isNewProject} rows={2} />
            )}
          />
        </Form.Item>

        {/* Client Name */}
        <Form.Item
          label="Client Name / Specification"
          validateStatus={errors.clientName ? "error" : ""}
          help={errors.clientName?.message}
        >
          <Controller
            name="clientName"
            control={control}
            render={({ field }) => (
              <Input
                {...field}
                disabled={!isNewProject}
                placeholder="Enter client name"
              />
            )}
          />
        </Form.Item>

        {/* Location */}
        <Form.Item
          label="Project Location"
          validateStatus={errors.projectLocation ? "error" : ""}
          help={errors.projectLocation?.message}
        >
          <Controller
            name="projectLocation"
            control={control}
            render={({ field }) => (
              <Input
                {...field}
                disabled={!isNewProject}
                placeholder="Enter location"
              />
            )}
          />
        </Form.Item>

        {/* Project Value */}
        <Form.Item
          label="Project Value (in crores)"
          validateStatus={errors.projectValue ? "error" : ""}
          help={errors.projectValue?.message}
        >
          <Controller
            name="projectValue"
            control={control}
            render={({ field }) => (
              <InputNumber
                {...field}
                disabled={!isNewProject}
                min={0}
                step={0.01}
                style={{ width: "100%" }}
                placeholder="Enter value in crores"
              />
            )}
          />
        </Form.Item>

        {/* Start Date */}
        <Form.Item
          label="Start Date"
          validateStatus={errors.startDate ? "error" : ""}
          help={errors.startDate?.message}
        >
          <Controller
            name="startDate"
            control={control}
            render={({ field }) => (
              <DatePicker
                {...field}
                value={field.value ? dayjs(field.value) : null}
                onChange={(date) => field.onChange(date)}
                style={{ width: "100%" }}
                disabledDate={(current) => {
                  return current && current <= dayjs().startOf("day");
                }}
              />
            )}
          />
        </Form.Item>

        {/* End Date */}
        <Form.Item
          label="End Date"
          validateStatus={errors.endDate ? "error" : ""}
          help={errors.endDate?.message}
        >
          <Controller
            name="endDate"
            control={control}
            render={({ field }) => (
              <DatePicker
                {...field}
                value={field.value ? dayjs(field.value) : null}
                onChange={(date) => field.onChange(date)}
                style={{ width: "100%" }}
                disabledDate={(current) => {
                  const start = control._formValues.startDate;
                  if (!start) {
                    return current && current <= dayjs().startOf("day");
                  }
                  return (
                    current &&
                    (current <= dayjs().startOf("day") ||
                      current <= dayjs(start).startOf("day"))
                  );
                }}
              />
            )}
          />
        </Form.Item>

        {/* Concrete Qty */}
        <Form.Item
          label="Concrete Quantity"
          validateStatus={errors.concreteQty ? "error" : ""}
          help={errors.concreteQty?.message}
        >
          <Controller
            name="concreteQty"
            control={control}
            render={({ field }) => (
              <InputNumber
                {...field}
                disabled={!isNewProject}
                min={0}
                precision={0}
                style={{ width: "100%" }}
                placeholder="Enter quantity"
              />
            )}
          />
        </Form.Item>

        {/* Fuel Cost */}
        <Form.Item
          label="Fuel Cost per Litre"
          validateStatus={errors.fuelCost ? "error" : ""}
          help={errors.fuelCost?.message}
        >
          <Controller
            name="fuelCost"
            control={control}
            render={({ field }) => (
              <InputNumber
              disabled={!isNewProject}
                {...field}
                min={0}
                step={0.01}
                style={{ width: "100%" }}
                placeholder="Enter fuel cost"
              />
            )}
          />
        </Form.Item>

        {/* Power Cost */}
        <Form.Item
          label="Power Cost per Unit"
          validateStatus={errors.powerCost ? "error" : ""}
          help={errors.powerCost?.message}
        >
          <Controller
            name="powerCost"
            control={control}
            render={({ field }) => (
              <InputNumber
                {...field}
                disabled={!isNewProject}
                min={0}
                step={0.01}
                style={{ width: "100%" }}
                placeholder="Enter power cost"
              />
            )}
          />
        </Form.Item>

        {/* File Path */}
        <Form.Item
          label="Name & Location to save workbook"
          validateStatus={errors.filePath ? "error" : ""}
          help={errors.filePath?.message}
        >
          <Controller
            name="filePath"
            control={control}
            render={({ field }) => (
              <div style={{ display: "flex", gap: "8px" }}>
                <div style={{ flex: 1 }}>
                  <Input
                    disabled={!isNewProject}
                    type="text"
                    placeholder="C:/path/to/workbook.xlsx"
                    {...field}
                    onBlur={(e) => {
                      field.onBlur();
                      const value = e.target.value;
                      if (value && !/\.(xlsx|xls)$/i.test(value)) {
                        message.error(
                          "Please enter a valid Excel file name (.xlsx or .xls)"
                        );
                      }
                    }}
                  />
                  <p style={{ color: "red" }}>{errors.filePath?.message}</p>
                </div>

                <Button
                  icon={<UploadOutlined />}
                  disabled={!isNewProject}
                  className="browse-btn"
                  onClick={async () => {
                    try {
                      // @ts-expect-error not typed yet
                      const dirHandle = await window.showDirectoryPicker();
                      field.onChange(dirHandle.name + "/");
                      openNotification("top", "Folder selected successfully!", "Success", "success");
                    } catch (error: unknown) {
                      if (error instanceof Error) {
                        openNotification("top", error.message, "Error", "error");
                      } else {
                        openNotification("top", "An unexpected error occurred", "Error", "error");
                      }
                    }
                  }}
                  >
                  Browse
                </Button>
              </div>
            )}
          />
        </Form.Item>

        {/* Submit */}
        <Form.Item>
          <Button type="primary" htmlType="submit" className="submit-btn">
            Save & Proceed
          </Button>
        </Form.Item>
      </Form>
    </div>
  );
}
