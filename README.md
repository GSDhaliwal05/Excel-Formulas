# Excel-Formulas
Project Overview

This tutorial demonstrates how to create a dynamic Excel invoice with:

Auto-calculation of totals and taxes.

Quantity conversion from kg to pounds.

Unit price auto-fill and total calculation per row.

Dropdown lists for products sourced from another sheet.

Automation via macros (VBA).

This project is ideal for beginners who want to create a functional invoice system in Excel.

1. Creating a New Excel File

Open Microsoft Excel.

Click File → New → Blank Workbook.

Save it immediately:

File → Save As → Choose location → Name: Invoice.xlsx

Save as type: Excel Workbook (*.xlsx)

Click Save

You now have a blank Excel workbook to start building your invoice.

2. Setting Up Sheets

Invoice Sheet:

Rename Sheet1 to Invoice.

Product Sheet:

Add a new sheet (+) and rename it Products.

This sheet will contain the product list for dropdowns.

3. Adding Product Names on Sheet2 (Products)

Go to Sheet2 (Products).

Enter your product names in column A, e.g.:

Product A
Product B
Product C
Product D


Optional: Create a named range:

Select the cells → Name Box (top-left) → Enter ProductsList.

Named ranges make dropdown references easier to manage.

4. Setting Up Invoice Table on Sheet1 (Invoice)

Start your table at row 19 (or any row you prefer).

Column	Purpose
B19:B30	Product Description (dropdown from Sheet2)
M19:M30	Quantity (kg) — auto-converts to pounds
O19:O30	Unit Price (default 4.49)
P19:P30	Total (=M*O/100)

Add totals below the table:

Cell	Formula / Label
E31	Subtotal: =SUM(P19:P30)
E32	Tax (13%): =E31*0.13
E33	Grand Total: =E31+E32
5. Adding Dropdowns for Products

Select B19:B30.

Go to Data → Data Validation → Data Validation.

Under Allow, select List.

In Source, enter:

=ProductsList


Or, if you didn’t name the range:

=Products!$A$1:$A$10


Click OK.

Alternate access: Right-click → Format Cells → Data Validation.

6. Adding the Macro for Automation

Open the VBA editor:

Method 1: Alt + F11

Method 2: Developer tab → Visual Basic

If Developer tab is not visible: File → Options → Customize Ribbon → Check Developer.

In the VBA editor:

Double-click Sheet1 (Invoice).

Paste the following code:

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


Save the workbook as a macro-enabled file:

File → Save As → Excel Macro-Enabled Workbook (*.xlsm)

7. How to Use the Invoice

Select a product from B19:B30 dropdown.

Enter a quantity in M19:M30 (kg).

Auto-converts to pounds.

Unit price in O is set to 4.49.

Total in P is calculated automatically.

Edit O19:O30 if needed — total recalculates automatically.

Subtotal, Tax, and Grand Total update automatically.

8. Alternate Access Methods

VBA Editor: Alt + F11 or Developer → Visual Basic

Macros: Developer → Macros → Run, Edit, or Create New

Data Validation: Data → Data Validation or Right-click → Format Cells → Data Validation

9. Visual Diagrams

Sheet1 (Invoice)

+---------+-------------------+------------+----------+
| B19:B30 | Product Description | M19:M30  | Quantity |
| O19:O30 | Unit Price         | P19:P30  | Total    |
+---------+-------------------+------------+----------+


Sheet2 (Products)

+---------+
| A1:A10  |
| Product A|
| Product B|
| Product C|
| Product D|
+---------+


The B column on Sheet1 references Sheet2 for the dropdown.

10. Notes / Tips for Beginners

Enable macros for automation: click Enable Content when opening .xlsm file.

Always save as Excel Macro-Enabled Workbook (.xlsm) if you want macros to run.

Adding more products: Add items to Sheet2 and update named range (ProductsList).

Adding more rows: Extend invoice ranges in Sheet1 and update macro ranges (M19:M30, O19:O30, P19:P30).

Editing formulas: Subtotal = SUM(P19:P30); Tax = Subtotal*0.13; Grand Total = Subtotal+Tax.

Accessing Developer tools: Alt + F11 or Developer tab → Visual Basic.

Feel Free to tune the VBS Code to your liking.

Thank You!
