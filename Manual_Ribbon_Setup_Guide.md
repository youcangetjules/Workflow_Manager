# Manual Ribbon Setup Guide for BASH Flow

## 🎯 **Why Manual Setup?**

Since Outlook requires ribbon customizations to be exported from Outlook itself (not created externally), we'll create the ribbon manually and then export it.

---

## 🚀 **Step-by-Step Instructions**

### **Step 1: Open Ribbon Customization**
1. **Open Outlook**
2. **Click:** `File` → `Options`
3. **Click:** `Customize Ribbon` (in the left panel)

### **Step 2: Create the BASH Flow Tab**
1. **Click:** `New Tab` (at the bottom of the right panel)
2. **Right-click** on the new tab that appears
3. **Select:** `Rename`
4. **Enter:** `BASH Flow`
5. **Click:** `OK`

### **Step 3: Create the Dashboard Group**
1. **Click:** `New Group` (under the BASH Flow tab)
2. **Right-click** on the new group
3. **Select:** `Rename`
4. **Enter:** `Dashboard`
5. **Click:** `OK`

#### **Add Dashboard Buttons:**
1. **Select:** `Dashboard` group
2. **From left panel:** Choose `Macros` from dropdown
3. **Add these buttons:**
   - **Button 1:**
     - **Name:** `Show Dashboard`
     - **Macro:** `ShowDashboardRibbon`
   - **Button 2:**
     - **Name:** `Refresh`
     - **Macro:** `RefreshDashboardRibbon`

### **Step 4: Create the Workflow Group**
1. **Click:** `New Group` (under the BASH Flow tab)
2. **Right-click** on the new group
3. **Select:** `Rename`
4. **Enter:** `Workflow`
5. **Click:** `OK`

#### **Add Workflow Buttons:**
1. **Select:** `Workflow` group
2. **Add these buttons:**
   - **Button 1:**
     - **Name:** `Process New`
     - **Macro:** `ProcessNewWorkflowItemsRibbon`
   - **Button 2:**
     - **Name:** `View Status`
     - **Macro:** `ViewAllRecordsRibbon`
   - **Button 3:**
     - **Name:** `Export`
     - **Macro:** `ExportDashboardDataRibbon`

### **Step 5: Create the Sandbox Group**
1. **Click:** `New Group` (under the BASH Flow tab)
2. **Right-click** on the new group
3. **Select:** `Rename`
4. **Enter:** `Sandbox`
5. **Click:** `OK`

#### **Add Sandbox Buttons:**
1. **Select:** `Sandbox` group
2. **Add these buttons:**
   - **Button 1:**
     - **Name:** `Initialize Sandbox`
     - **Macro:** `InitializeSandboxRibbon`
   - **Button 2:**
     - **Name:** `Fix Database`
     - **Macro:** `FixDatabaseTableRibbon`
   - **Button 3:**
     - **Name:** `Check Database`
     - **Macro:** `CheckDatabaseStructureRibbon`

### **Step 6: Create the Diagnostics Group**
1. **Click:** `New Group` (under the BASH Flow tab)
2. **Right-click** on the new group
3. **Select:** `Rename`
4. **Enter:** `Diagnostics`
5. **Click:** `OK`

#### **Add Diagnostics Buttons:**
1. **Select:** `Diagnostics` group
2. **Add these buttons:**
   - **Button 1:**
     - **Name:** `Run Diagnostics`
     - **Macro:** `RunSandboxDiagnosticsRibbon`
   - **Button 2:**
     - **Name:** `View Logs`
     - **Macro:** `ViewSandboxLogsRibbon`

### **Step 7: Export the Customization**
1. **Click:** `Import/Export` button (at the bottom)
2. **Select:** `Export all customizations`
3. **Navigate to:** `C:\Lumen\Workflow Manager\`
4. **File name:** `BASH_Flow_Ribbon.exportedUI`
5. **Click:** `Save`

### **Step 8: Test the Ribbon**
1. **Click:** `OK` to close the Options dialog
2. **Restart Outlook**
3. **Look for:** "BASH Flow" tab in the ribbon
4. **Test each button** to ensure they work

---

## 📋 **Macro Reference List**

Use these exact macro names when assigning buttons:

### **Dashboard Functions:**
- `ShowDashboardRibbon`
- `RefreshDashboardRibbon`

### **Workflow Functions:**
- `ProcessNewWorkflowItemsRibbon`
- `ViewAllRecordsRibbon`
- `ExportDashboardDataRibbon`

### **Sandbox Functions:**
- `InitializeSandboxRibbon`
- `FixDatabaseTableRibbon`
- `CheckDatabaseStructureRibbon`

### **Diagnostic Functions:**
- `RunSandboxDiagnosticsRibbon`
- `ViewSandboxLogsRibbon`

---

## 🎯 **Expected Result**

After setup, you should see:

```
┌─ BASH Flow Tab ─────────────────────────────────────┐
│ ┌─ Dashboard ─┐ ┌─ Workflow ─┐ ┌─ Sandbox ─┐ ┌─ Diagnostics ─┐ │
│ │ Show        │ │ Process    │ │ Initialize │ │ Run           │ │
│ │ Dashboard   │ │ New        │ │ Sandbox    │ │ Diagnostics   │ │
│ │             │ │            │ │            │ │               │ │
│ │ Refresh     │ │ View       │ │ Fix        │ │ View          │ │
│ │             │ │ Status     │ │ Database   │ │ Logs          │ │
│ │             │ │            │ │            │ │               │ │
│ │             │ │ Export     │ │ Check      │ │               │ │
│ │             │ │            │ │ Database   │ │               │ │
└────────────────────────────────────────────────────┘
```

---

## 🚨 **Troubleshooting**

### **"Macro not found" errors:**
1. **Ensure all VBA modules are imported**
2. **Run:** `SetupBASHFlowModules()` first
3. **Verify:** Module names match exactly

### **"Button not working":**
1. **Check:** Macro assignment is correct
2. **Verify:** Function exists in imported modules
3. **Test:** Run function directly in VBA Editor

### **"Ribbon tab not showing":**
1. **Check:** Outlook was restarted after setup
2. **Verify:** Ribbon customization is enabled
3. **Try:** `File` → `Options` → `Advanced` → `Customize Ribbon`

---

## ✅ **Success Checklist**

- [ ] BASH Flow tab appears in Outlook ribbon
- [ ] 4 groups are visible (Dashboard, Workflow, Sandbox, Diagnostics)
- [ ] All buttons are present and named correctly
- [ ] Each button is assigned to the correct macro
- [ ] Buttons respond when clicked
- [ ] Dashboard opens from "Show Dashboard" button
- [ ] No error messages when using buttons

---

## 🔄 **Sharing the Ribbon**

Once you've created and exported the ribbon:

1. **The `BASH_Flow_Ribbon.exportedUI` file** can be shared with others
2. **They can import it** using `Import/Export` → `Import customization file`
3. **They need the VBA modules** imported first for the buttons to work

This manual method ensures compatibility with Outlook's ribbon system! 🎯
