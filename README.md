
# 📦 Excel Invoice System By Guransh Dhaliwal

## 🧾 Overview

This project walks you through building a dynamic invoice system in Microsoft Excel using formulas, dropdowns, and VBA macros. It’s designed for beginners who want to automate calculations, manage products from a separate sheet, and convert quantities from kilograms to pounds.

---

## 🚀 Features

- Auto-calculation of totals, taxes, and grand total  
- Quantity conversion from kilograms to pounds  
- Product selection via dropdown (linked to a separate sheet)  
- Unit price autofill and row-wise total calculation  
- VBA macro for automation and dynamic updates  
- Named ranges for easy data validation  
- Expandable design for more products and rows  
- Developer shortcuts for quick access to tools

---

## 📁 Setup Instructions

### 1. Create Your Workbook

- Open Excel → **File → New → Blank Workbook**
- Save immediately:
  - File → Save As → Choose location
  - Name: `Invoice.xlsm`
  - Save as type: **Excel Macro-Enabled Workbook (*.xlsm)**

---

## 📑 Setting Up Sheets

### Sheet1 → Rename to `Invoice`  
This is your main invoice interface.

### Sheet2 → Rename to `Products`  
This sheet stores your product list for dropdowns.

---

## 🛒 Adding Product Names on Sheet2 (Products)

Enter product names in column A:

```
Product A  
Product B  
Product C  
Product D
```

Optional: Create a named range:

- Select cells A1:A10  
- In the Name Box (top-left), type: `ProductsList`  
- Press Enter

---

## 📊 Invoice Table Structure (Sheet1: Invoice)

Start your table at row 19:

| Column   | Description                       |
|----------|-----------------------------------|
| B19:B30  | Product Description (dropdown)    |
| M19:M30  | Quantity (kg) — auto-converts     |
| O19:O30  | Unit Price (default: 4.49)        |
| P19:P30  | Total (`=M*O/100`)                |

### Totals Section

| Cell | Formula / Label            |
|------|----------------------------|
| E31  | Subtotal: `=SUM(P19:P30)` |
| E32  | Tax (13%): `=E31*0.13`     |
| E33  | Grand Total: `=E31+E32`   |

---

## 🔽 Adding Dropdowns for Products

1. Select **B19:B30**  
2. Go to **Data → Data Validation → Data Validation**  
3. Under **Allow**, select **List**  
4. In Source, enter:

```excel
=ProductsList
```

Or, if you didn’t name the range:

```excel
=Products!$A$1:$A$10
```

5. Click **OK**

Alternate access:  
Right-click → Format Cells → Data Validation

---

## 🧠 VBA Macro for Automation

### Open the VBA editor:

- Method 1: Press `Alt + F11`  
- Method 2: Developer tab → Visual Basic  
  - If Developer tab is hidden:  
    - File → Options → Customize Ribbon → Check **Developer**

### Paste the following code into Sheet1 (Invoice):

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngM As Range, rngO As Range
    Dim row As Long
    Dim mValue As Double
    Dim pFormula As String

    If Target.CountLarge > 1 Then Exit Sub

    Set rngM = Me.Range("M19:M30")
    Set rngO = Me.Range("O19:O30")

    Application.EnableEvents = False

    If Not Intersect(Target, rngM) Is Nothing Then
        row = Target.Row
        If IsNumeric(Target.Value) And Target.Value <> "" Then
            mValue = Target.Value * 2.20462262
            Target.Value = mValue
        Else
            Target.Value = ""
        End If
        Me.Cells(row, "O").Value = 4.49
        pFormula = "=M" & row & "*O" & row & "/100"
        Me.Cells(row, "P").Formula = pFormula
    End If

    If Not Intersect(Target, rngO) Is Nothing Then
        row = Target.Row
        pFormula = "=M" & row & "*O" & row & "/100"
        Me.Cells(row, "P").Formula = pFormula
    End If

    Application.EnableEvents = True
End Sub
```

💾 Save your workbook as `.xlsm` to enable macros.

---

## 🧪 How to Use the Invoice

- Select a product from the dropdown in **B19:B30**
- Enter quantity in **M19:M30** (kg) — auto-converts to pounds
- Unit price autofills in **O19:O30**
- Total in **P19:P30** calculates automatically
- Subtotal, Tax, and Grand Total update in real time
- You can manually adjust unit prices — totals will recalculate

---

## 🧠 Tips for Beginners

- Enable macros: Click **Enable Content** when opening the `.xlsm` file  
- Add more products: Update `Products` sheet and named range  
- Add more rows: Extend ranges in VBA and formulas  
- Formula references:
  - Subtotal: `=SUM(P19:P30)`
  - Tax: `=Subtotal * 0.13`
  - Grand Total: `=Subtotal + Tax`
- Access Developer tools: `Alt + F11` or **Developer tab → Visual Basic**

---

## 🧭 Developer Shortcuts

| Action              | Shortcut / Location                     |
|---------------------|------------------------------------------|
| Open VBA Editor     | `Alt + F11` or Developer → Visual Basic |
| Run/Edit Macros     | Developer → Macros                      |
| Data Validation     | Data → Data Validation or Right-click   |

---

## 🧱 Visual Layout

### Sheet1 (Invoice)

```
+---------+-------------------+------------+----------+
| B19:B30 | Product Description| M19:M30   | Quantity |
| O19:O30 | Unit Price         | P19:P30   | Total    |
+---------+-------------------+------------+----------+
```

### Sheet2 (Products)

```
+---------+
| A1:A10  |
| Product A|
| Product B|
| Product C|
| Product D|
+---------+
```


