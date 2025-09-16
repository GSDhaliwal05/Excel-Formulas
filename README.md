# Excel-Formulas

## Project Overview

This tutorial demonstrates how to create a dynamic Excel invoice with:

- Auto-calculation of totals and taxes
- Quantity conversion from kg to pounds
- Unit price auto-fill and total calculation per row
- Dropdown lists for products sourced from another sheet
- Automation via macros (VBA)

This project is ideal for beginners who want to create a functional invoice system in Excel.

## Creating a New Excel File

1. Open Microsoft Excel.
2. Click **File → New → Blank Workbook**.
3. Save it immediately:
   - File → Save As → Choose location → Name: `Invoice.xlsx`
   - Save as type: Excel Workbook (*.xlsx)
   - Click Save

You now have a blank Excel workbook to start building your invoice.

## Setting Up Sheets

### Invoice Sheet
Rename Sheet1 to `Invoice`.

### Product Sheet
Add a new sheet (+) and rename it `Products`.  
This sheet will contain the product list for dropdowns.

## Adding Product Names on Sheet2 (Products)

Enter your product names in column A:

Product A
Product B
Product C
Product D

sql
Copy code

Optional: Create a named range:

- Select the cells → Name Box (top-left) → Enter `ProductsList`

Named ranges make dropdown references easier to manage.

## Setting Up Invoice Table on Sheet1 (Invoice)

Start your table at row 19 (or any row you prefer):

| Column   | Purpose                        |
|----------|--------------------------------|
| B19:B30  | Product Description (dropdown)  |
| M19:M30  | Quantity (kg) — auto-converts  |
| O19:O30  | Unit Price (default 4.49)      |
| P19:P30  | Total (`=M*O/100`)             |

Add totals below the table:

| Cell | Formula / Label          |
|------|--------------------------|
| E31  | Subtotal: `=SUM(P19:P30)`|
| E32  | Tax (13%): `=E31*0.13`  |
| E33  | Grand Total: `=E31+E32` |

## Adding Dropdowns for Products

1. Select B19:B30
2. Go to **Data → Data Validation → Data Validation**
3. Under **Allow**, select **List**
4. In Source, enter:

=ProductsList

pgsql
Copy code

Or, if you didn’t name the range:

=Products!$A$1:$A$10

mathematica
Copy code

5. Click OK

Alternate access: Right-click → Format Cells → Data Validation

## Adding the Macro for Automation

Open the VBA editor:

- Method 1: Alt + F11  
- Method 2: Developer tab → Visual Basic  
  - If Developer tab is not visible: File → Options → Customize Ribbon → Check Developer

Paste the following code in Sheet1 (Invoice):

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
Save the workbook as a macro-enabled file: .xlsm

How to Use the Invoice
Select a product from B19:B30 dropdown.

Enter a quantity in M19:M30 (kg)

Auto-converts to pounds

Unit price in O is set to 4.49

Total in P is calculated automatically

Edit O19:O30 if needed — total recalculates automatically

Subtotal, Tax, and Grand Total update automatically

Alternate Access Methods
VBA Editor: Alt + F11 or Developer → Visual Basic

Macros: Developer → Macros → Run, Edit, or Create New

Data Validation: Data → Data Validation or Right-click → Format Cells → Data Validation

Visual Diagrams
Sheet1 (Invoice)

mathematica
Copy code
+---------+-------------------+------------+----------+
| B19:B30 | Product Description | M19:M30  | Quantity |
| O19:O30 | Unit Price         | P19:P30  | Total    |
+---------+-------------------+------------+----------+
Sheet2 (Products)

mathematica
Copy code
+---------+
| A1:A10  |
| Product A|
| Product B|
| Product C|
| Product D|
+---------+
#Notes / Tips for Beginners
Enable macros for automation: click Enable Content when opening .xlsm file.

Always save as Excel Macro-Enabled Workbook (.xlsm)

Adding more products: Add items to Sheet2 and update named range (ProductsList)

Adding more rows: Extend invoice ranges in Sheet1 and update macro ranges (M19:M30, O19:O30, P19:P30)

Editing formulas: Subtotal = SUM(P19:P30); Tax = Subtotal*0.13; Grand Total = Subtotal+Tax

Accessing Developer tools: Alt + F11 or Developer tab → Visual Basic

Feel free to tweak the VBA code as needed.

Thank you!
