# SharePoint CRUD WebPart (Internship Task)

## 📌 Overview  
This project is a **SharePoint Framework (SPFx) web part** built during my internship.  
It demonstrates how to:  
- Create and manage a **SharePoint list** (CRUD operations).  
- Use **Fluent UI** for modern UI components.  
- Implement **People Picker** and list operations using **PnP JS**.  
- Apply **React fundamentals** – `state`, `props`, and parent ↔ child communication.

---

## ⚙️ Features  

- **SharePoint List Integration**  
  - Create, Read, Update, and Delete (CRUD) items from a SharePoint list.  
  - Example names used: `Abhay`, `Shivam`, `Abhishek`.  

- **Fluent UI Components**  
  - User-friendly and modern UI elements.  
  - Example: buttons, inputs, dialogs.  

- **PnP JS Integration**  
  - Used for **People Picker**.  
  - Used for efficient SharePoint list operations.  

- **React Concepts**  
  - **State Management**: Used `useState` (functional components) to track list items.  
  - **Props**: Passed data from parent → child components.  
  - **Child-to-Parent Communication**: Used callback functions so child components can update parent state.  
  - Covered both **functional props** and **class state** understanding.  

---

## 🏗️ Project Structure  

```
src/webparts/
├── components/
│ ├── EmployeeCrud.tsx # Parent component (main list state + CRUD logic)
│ ├── IEmployeeCrudProps.tsx # Child component (adds new items via props)
│ ├── EmployeeCrud.Module.scss # Default scss file.
│ └── ListItem.tsx # List rendering & delete/update
└── /loc
  └── EmployeeCrudWebPart.ts # Webparts render controller.
  └── pnpConfig.ts # Created for pnp import and avoid version conflict.

```
  
---

## 🚀 How to Run  

1. Clone the repo  
   ```bash
   git clone https://github.com/Vipul1020/sharepoint-crud-webpart.git
   cd sharepoint-crud-webpart

2. Install Dependencies
   ```
   npm install
   
   ```

4. Start local workbench - gulp serve

5. Test on SharePoint Workbench - https://yourtenant.sharepoint.com/_layouts/15/workbench.aspx

---

## Tech Stack  

- SharePoint Framework (SPFx)
- React + TypeScript
- Fluent UI
- PnP JS

---

## 🎯 Learning Outcomes

- Gained practical experience with SPFx web part development.

- Learned to integrate PnP JS for SharePoint operations.

- Strengthened React fundamentals: state, props, and component communication.

- Hands-on CRUD implementation with SharePoint lists.



  
