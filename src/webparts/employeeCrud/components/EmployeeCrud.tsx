import * as React from "react";
import {
  TextField,
  PrimaryButton,
  DefaultButton,
  DetailsList,
  IColumn,
  Stack,
  Persona,
  PersonaSize,
  DatePicker,
  DayOfWeek,
  Label,
} from "@fluentui/react";
import { ShimmeredDetailsList } from "@fluentui/react/lib/ShimmeredDetailsList";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { sp } from "../pnpConfig";
import { IEmployeeCrudProps } from "./IEmployeeCrudProps";

type SPUserField = {
  Id: number;
  Title: string;
  EMail?: string;
};

type EmployeeItem = {
  Id: number;
  Title: string;
  Department?: string;
  Email?: string;
  EmployeeID?: string;
  DateOfJoining?: string;
  // person field
  PeoplepICKER?: SPUserField;
};

// === CHANGE THIS to your Person column internal name ===
const PERSON_FIELD_INTERNAL_NAME = "PeoplepICKER";
// ======================================================

const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const generateEmployeeId = (): string => `EMP-${Date.now()}`;

const EmployeeCrud: React.FC<IEmployeeCrudProps> = (props) => {
  const [name, setName] = React.useState("");
  const [department, setDepartment] = React.useState("");
  const [email, setEmail] = React.useState("");
  const [dateOfJoining, setDateOfJoining] = React.useState<Date | null>(null);
  const [employeeId, setEmployeeId] = React.useState<string>("");
  const [employees, setEmployees] = React.useState<EmployeeItem[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [editingId, setEditingId] = React.useState<number | null>(null);
  const [editDepartment, setEditDepartment] = React.useState("");
  const [selectedPeople, setSelectedPeople] = React.useState<any[]>([]);

  const [lookupId, setLookupId] = React.useState("");
  const [lookupResult, setLookupResult] = React.useState<EmployeeItem | null>(
    null
  );

  const peoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    spHttpClient: props.context.spHttpClient,
    msGraphClientFactory: props.context.msGraphClientFactory,
  };

  // Helper: extract an identifier (loginName/email/key) from PeoplePicker item
  const extractUserIdentifier = (pickerItem: any): string | null => {
    if (!pickerItem) return null;
    return (
      pickerItem.loginName ||
      pickerItem.key ||
      pickerItem.id ||
      pickerItem.secondaryText || // usually email
      pickerItem.email ||
      pickerItem.text ||
      null
    );
  };

  // Ensure SP user exists and return SP user id
  const ensureSpUserId = async (
    userIdentifier: string
  ): Promise<number | null> => {
    if (!userIdentifier) return null;
    try {
      const ensured = await sp.web.ensureUser(userIdentifier);
      // pnpjs shape varies by version; try common locations
      const spUserId =
        (ensured && (ensured as any).data && (ensured as any).data.Id) ||
        (ensured && (ensured as any).Id) ||
        null;
      return spUserId;
    } catch (err) {
      console.error("ensureUser failed for", userIdentifier, err);
      return null;
    }
  };

  // Load items — include person field via select + expand
  const loadItems = React.useCallback(async () => {
    setLoading(true);
    try {
      const selectFields = [
        "Id",
        "Title",
        "Department",
        "Email",
        "EmployeeID",
        "DateOfJoining",
        `${PERSON_FIELD_INTERNAL_NAME}/Id`,
        `${PERSON_FIELD_INTERNAL_NAME}/Title`,
        `${PERSON_FIELD_INTERNAL_NAME}/EMail`,
      ];
      const expandFields = [PERSON_FIELD_INTERNAL_NAME];

      const items: EmployeeItem[] = await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.select(selectFields.join(","))
        .expand(expandFields.join(","))
        .get();

      setEmployees(items);
    } catch (err) {
      console.error("Error loading items:", err);
      alert("Failed to load employees. See console for details.");
    } finally {
      setLoading(false);
    }
  }, []);

  React.useEffect(() => {
    loadItems();
  }, [loadItems]);

  const resetForm = () => {
    setName("");
    setDepartment("");
    setEmail("");
    setDateOfJoining(null);
    setEmployeeId("");
    setSelectedPeople([]);
  };

  const addItem = async () => {
    if (!name.trim() || !department.trim() || !email.trim()) {
      alert("Please fill all required fields.");
      return;
    }
    if (!emailRegex.test(email)) {
      alert("Please enter a valid email address.");
      return;
    }

    setSaving(true);
    try {
      const empIdToStore = employeeId?.trim() || generateEmployeeId();
      const dojIso = dateOfJoining ? dateOfJoining.toISOString() : null;

      const payload: any = {
        Title: name.trim(),
        Department: department.trim(),
        Email: email.trim(),
        EmployeeID: empIdToStore,
        DateOfJoining: dojIso,
      };

      // If a person was selected, resolve to SP user id and set <FieldInternalName>Id
      if (selectedPeople && selectedPeople.length > 0) {
        const identifier = extractUserIdentifier(selectedPeople[0]);
        const spUserId = await ensureSpUserId(identifier || "");
        if (spUserId) {
          payload[`${PERSON_FIELD_INTERNAL_NAME}Id`] = spUserId;
        } else {
          console.warn(
            "Could not resolve selected person to a SharePoint user id. Skipping person field."
          );
        }
      }

      await sp.web.lists.getByTitle("EmployeeDetails").items.add(payload);

      resetForm();
      await loadItems();
    } catch (err) {
      console.error("Error adding item:", err);
      alert("Failed to add employee. See console for details.");
    } finally {
      setSaving(false);
    }
  };

  const startEdit = (item: EmployeeItem) => {
    setEditingId(item.Id);
    setEditDepartment(item.Department || "");
  };

  const saveEdit = async (id: number) => {
    if (!editDepartment.trim()) {
      alert("Department cannot be empty.");
      return;
    }
    setSaving(true);
    try {
      await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.getById(id)
        .update({
          Department: editDepartment.trim(),
        });
      setEditingId(null);
      await loadItems();
    } catch (err) {
      console.error("Error updating item:", err);
      alert("Failed to update employee. See console for details.");
    } finally {
      setSaving(false);
    }
  };

  const cancelEdit = () => {
    setEditingId(null);
    setEditDepartment("");
  };

  const deleteItem = async (id: number) => {
    if (!confirm("Are you sure you want to delete this employee?")) return;
    setSaving(true);
    try {
      await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.getById(id)
        .delete();
      await loadItems();
    } catch (err) {
      console.error("Error deleting item:", err);
      alert("Failed to delete employee. See console for details.");
    } finally {
      setSaving(false);
    }
  };

  const fetchByEmployeeId = async (idValue: string) => {
    if (!idValue.trim()) {
      alert("Enter EmployeeID to lookup.");
      return;
    }
    setLoading(true);
    setLookupResult(null);
    try {
      // filter by EmployeeID and select/expand person field
      const results: EmployeeItem[] = await sp.web.lists
        .getByTitle("EmployeeDetails")
        .items.filter(`EmployeeID eq '${idValue.trim()}'`)
        .select(
          "Id,Title,Department,Email,EmployeeID,DateOfJoining," +
            `${PERSON_FIELD_INTERNAL_NAME}/Id,${PERSON_FIELD_INTERNAL_NAME}/Title,${PERSON_FIELD_INTERNAL_NAME}/EMail`
        )
        .expand(PERSON_FIELD_INTERNAL_NAME)
        .top(1)
        .get();

      if (results && results.length) {
        setLookupResult(results[0]);
      } else {
        alert("No employee found with that EmployeeID.");
        setLookupResult(null);
      }
    } catch (err) {
      console.error("Error fetching by EmployeeID:", err);
      alert("Lookup failed. See console for details.");
    } finally {
      setLoading(false);
    }
  };

  const columns: IColumn[] = [
    {
      key: "empid",
      name: "Emp ID",
      fieldName: "EmployeeID",
      minWidth: 100,
      onRender: (item: EmployeeItem) => <strong>{item.EmployeeID}</strong>,
    },
    {
      key: "name",
      name: "Name",
      fieldName: "Title",
      minWidth: 140,
      onRender: (item: EmployeeItem) => (
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <Persona
            text={item.Title}
            size={PersonaSize.size32}
            secondaryText={item.Email}
          />
        </div>
      ),
    },
    {
      key: "dept",
      name: "Department",
      fieldName: "Department",
      minWidth: 140,
      onRender: (item: EmployeeItem) =>
        editingId === item.Id ? (
          <TextField
            value={editDepartment}
            onChange={(_, v) => setEditDepartment(v || "")}
          />
        ) : (
          <span>{item.Department}</span>
        ),
    },
    {
      key: "person",
      name: "Person",
      fieldName: PERSON_FIELD_INTERNAL_NAME,
      minWidth: 160,
      onRender: (item: EmployeeItem) =>
        item.PeoplepICKER ? (
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <Persona
              text={item.PeoplepICKER.Title}
              secondaryText={item.PeoplepICKER.EMail}
              size={PersonaSize.size32}
            />
          </div>
        ) : (
          <span>-</span>
        ),
    },
    {
      key: "doj",
      name: "Date of Joining",
      fieldName: "DateOfJoining",
      minWidth: 140,
      onRender: (item: EmployeeItem) =>
        item.DateOfJoining ? (
          new Date(item.DateOfJoining).toLocaleDateString()
        ) : (
          <span>-</span>
        ),
    },
    {
      key: "actions",
      name: "Actions",
      minWidth: 180,
      onRender: (item: EmployeeItem) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          {editingId === item.Id ? (
            <>
              <PrimaryButton
                text="Save"
                onClick={() => saveEdit(item.Id)}
                disabled={saving}
              />
              <DefaultButton
                text="Cancel"
                onClick={cancelEdit}
                disabled={saving}
              />
            </>
          ) : (
            <>
              <DefaultButton text="Edit" onClick={() => startEdit(item)} />
              <DefaultButton
                text="Delete"
                onClick={() => deleteItem(item.Id)}
                disabled={saving}
              />
            </>
          )}
        </Stack>
      ),
    },
  ];

  const onPeopleChange = (items: any[]) => {
    setSelectedPeople(items || []);
    if (items && items.length > 0) {
      const first = items[0];
      if (first.text && !name) setName(first.text);
      if ((first.secondaryText || first.email) && !email) {
        setEmail(first.secondaryText || first.email || "");
      }
    }
    console.log("PeoplePicker selected:", items);
  };

  return (
    <div style={{ padding: 20 }}>
      <h2>Employee CRUD — Premium</h2>

      <Stack tokens={{ childrenGap: 12 }}>
        <TextField
          label="Employee ID (optional)"
          value={employeeId}
          onChange={(_, val) => setEmployeeId(val || "")}
          placeholder="Leave blank to auto-generate"
        />

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

        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 12 }}>
          <div style={{ width: 260 }}>
            <Label>Date of Joining</Label>
            <DatePicker
              firstDayOfWeek={DayOfWeek.Monday}
              value={dateOfJoining || undefined}
              onSelectDate={(date) => setDateOfJoining(date || null)}
              placeholder="Select a date..."
            />
          </div>

          <div style={{ flex: 1 }}>
            <PeoplePicker
              context={peoplePickerContext as any}
              titleText="Select People (optional)"
              personSelectionLimit={1}
              showtooltip={true}
              required={false}
              defaultSelectedUsers={[]}
              onChange={onPeopleChange}
              principalTypes={[PrincipalType.User]}
              resolveDelay={500}
            />

            {selectedPeople && selectedPeople.length > 0 && (
              <div style={{ marginTop: 8 }}>
                <strong>Selected:</strong>{" "}
                {selectedPeople
                  .map((p, i) =>
                    p.text ? p.text : p.id || p.loginName || `user${i}`
                  )
                  .join(", ")}
              </div>
            )}
          </div>
        </Stack>

        <PrimaryButton
          text={saving ? "Saving..." : "Add Employee"}
          onClick={addItem}
          disabled={saving}
        />
      </Stack>

      <div style={{ marginTop: 28 }}>
        <h3>Lookup employee by EmployeeID</h3>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <TextField
            placeholder="Enter EmployeeID (e.g. EMP-1623456789012)"
            value={lookupId}
            onChange={(_, v) => setLookupId(v || "")}
          />
          <PrimaryButton
            text="Fetch"
            onClick={() => fetchByEmployeeId(lookupId)}
          />
          <DefaultButton
            text="Clear"
            onClick={() => {
              setLookupId("");
              setLookupResult(null);
            }}
          />
        </Stack>

        {lookupResult && (
          <div
            style={{
              marginTop: 12,
              padding: 12,
              border: "1px solid #ddd",
              borderRadius: 6,
            }}
          >
            <h4>Details for {lookupResult.EmployeeID}</h4>
            <p>
              <strong>Name:</strong> {lookupResult.Title}
            </p>
            <p>
              <strong>Email:</strong> {lookupResult.Email}
            </p>
            <p>
              <strong>Department:</strong> {lookupResult.Department}
            </p>
            <p>
              <strong>Person:</strong>{" "}
              {lookupResult.PeoplepICKER
                ? `${lookupResult.PeoplepICKER.Title} (${
                    lookupResult.PeoplepICKER.EMail || ""
                  })`
                : "-"}
            </p>
            <p>
              <strong>Date of Joining:</strong>{" "}
              {lookupResult.DateOfJoining
                ? new Date(lookupResult.DateOfJoining).toLocaleDateString()
                : "-"}
            </p>
            <p>
              <strong>SP Item Id:</strong> {lookupResult.Id}
            </p>
          </div>
        )}
      </div>

      <h3 style={{ marginTop: 24 }}>Employee List</h3>

      {loading ? (
        <ShimmeredDetailsList
          items={employees}
          columns={columns}
          selectionMode={0}
        />
      ) : (
        <DetailsList items={employees} columns={columns} selectionMode={0} />
      )}
    </div>
  );
};

export default EmployeeCrud;
