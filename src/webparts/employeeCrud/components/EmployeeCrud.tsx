import * as React from "react";
import {
  TextField,
  PrimaryButton,
  DetailsList,
  IColumn,
} from "@fluentui/react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "../pnpConfig";
import { IEmployeeCrudProps } from "./IEmployeeCrudProps";

const EmployeeCrud: React.FC<IEmployeeCrudProps> = (props) => {
  const [name, setName] = React.useState("");
  const [department, setDepartment] = React.useState("");
  const [email, setEmail] = React.useState("");
  const [employees, setEmployees] = React.useState<any[]>([]);

  React.useEffect(() => {
    loadItems();
  }, []);

  const loadItems = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.get();
      setEmployees(items);
    } catch (err) {
      console.error("Error loading items:", err);
    }
  };

  const addItem = async () => {
    if (!name || !department || !email) {
      alert("Please fill all fields");
      return;
    }
    try {
      await sp.web.lists.getByTitle("EmployeeDetails").items.add({
        Title: name,
        Department: department,
        Email: email,
      });
      setName("");
      setDepartment("");
      setEmail("");
      loadItems();
    } catch (err) {
      console.error("Error adding item:", err);
    }
  };

  const updateItem = async (id: number) => {
    try {
      await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.getById(id)
        .update({
          Department: "Updated Department",
        });
      loadItems();
    } catch (err) {
      console.error("Error updating item:", err);
    }
  };

  const deleteItem = async (id: number) => {
    try {
      await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.getById(id)
        .delete();
      loadItems();
    } catch (err) {
      console.error("Error deleting item:", err);
    }
  };

  const columns: IColumn[] = [
    { key: "name", name: "Name", fieldName: "Title", minWidth: 100 },
    { key: "dept", name: "Department", fieldName: "Department", minWidth: 100 },
    { key: "email", name: "Email", fieldName: "Email", minWidth: 150 },
    {
      key: "actions",
      name: "Actions",
      minWidth: 150,
      onRender: (item: any) => (
        <div>
          <PrimaryButton text="Update" onClick={() => updateItem(item.Id)} />
          <PrimaryButton
            text="Delete"
            style={{ marginLeft: 8 }}
            onClick={() => deleteItem(item.Id)}
          />
        </div>
      ),
    },
  ];

  return (
    <div style={{ padding: 20 }}>
      <h2>Employee CRUD Web Part</h2>

      <TextField
        label="Name"
        value={name}
        onChange={(_, val) => setName(val || "")}
      />
      <TextField
        label="Department"
        value={department}
        onChange={(_, val) => setDepartment(val || "")}
      />
      <TextField
        label="Email"
        value={email}
        onChange={(_, val) => setEmail(val || "")}
      />

      {/* People Picker */}
      <PeoplePicker
        context={{
          absoluteUrl: props.context.pageContext.web.absoluteUrl,
          msGraphClientFactory: props.context.msGraphClientFactory,
          spHttpClient: props.context.spHttpClient,
        }}
        titleText="Select People"
        personSelectionLimit={3}
        showtooltip={true}
        required={false}
        defaultSelectedUsers={[
          "Abhishek.Kumar001@366pidev.onmicrosoft.com",
          "Shivam",
          "Abhay",
        ]}
        onChange={(items) => console.log("Selected people:", items)}
        principalTypes={[PrincipalType.User]}
        resolveDelay={500}
      />

      <PrimaryButton
        text="Add Employee"
        onClick={addItem}
        style={{ marginTop: 12 }}
      />

      <h3 style={{ marginTop: 20 }}>Employee List</h3>
      <DetailsList items={employees} columns={columns} />
    </div>
  );
};

export default EmployeeCrud;
