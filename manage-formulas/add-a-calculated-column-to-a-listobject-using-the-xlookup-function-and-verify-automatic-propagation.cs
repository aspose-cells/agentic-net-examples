// Title: Add an XLOOKUP calculated column to an Aspose.Cells ListObject and confirm automatic row propagation in C#
// Common Searches: C# Aspose.Cells add calculated column with XLOOKUP to a table | how to use structured references in Aspose.Cells ListObject for XLOOKUP | verify that XLOOKUP formula propagates to all rows in an Aspose.Cells ListObject | create lookup table and related calculated column using Aspose.Cells for .NET
// Tags: add XLOOKUP calculated column Aspose.Cells | listobject structured reference formula C# | automatic formula propagation Excel table Aspose.Cells | lookup table to orders table Aspose.Cells | C# Aspose.Cells XLOOKUP example

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // Required for ListObject

// The sample creates a workbook with two tables—Table1 (ID, Name) and Orders (OrderID, CustomerID, CustomerName). It adds an XLOOKUP formula as a calculated column in the Orders ListObject to retrieve the customer name from Table1 using structured references. The code prints each cell's formula to demonstrate that the formula automatically propagates to every data row, then saves the file as CalculatedColumn_XLOOKUP.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // ---------- Create lookup table (Table1) ----------
            // Header
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            // Data
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");

            // Define Table1 covering A1:B4
            int lookupIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject lookupTable = sheet.ListObjects[lookupIdx];
            lookupTable.DisplayName = "Table1";

            // ---------- Create main data table (Orders) ----------
            // Header
            sheet.Cells["D1"].PutValue("OrderID");
            sheet.Cells["E1"].PutValue("CustomerID");
            sheet.Cells["F1"].PutValue("CustomerName"); // placeholder for calculated column
            // Data
            sheet.Cells["D2"].PutValue(100);
            sheet.Cells["E2"].PutValue(2);
            sheet.Cells["D3"].PutValue(101);
            sheet.Cells["E3"].PutValue(3);
            sheet.Cells["D4"].PutValue(102);
            sheet.Cells["E4"].PutValue(1);

            // Define Orders table covering D1:F4
            int ordersIdx = sheet.ListObjects.Add(0, 3, 3, 5, true);
            ListObject ordersTable = sheet.ListObjects[ordersIdx];
            ordersTable.DisplayName = "Orders";

            // ---------- Add calculated column using XLOOKUP ----------
            // XLOOKUP will retrieve the Name from Table1 based on CustomerID
            // Structured reference: [@CustomerID] refers to the current row's CustomerID
            string xlookupFormula = "=XLOOKUP([@CustomerID],Table1[ID],Table1[Name],\"Not Found\")";

            // The calculated column is the third column in the Orders table (index 2, zero‑based)
            ordersTable.ListColumns[2].Formula = xlookupFormula;

            // ---------- Verify automatic propagation ----------
            // Each cell in the calculated column should contain the same formula (adjusted per row)
            int dataRowCount = ordersTable.DataRange.RowCount; // number of data rows (excluding header)
            for (int i = 0; i < dataRowCount; i++)
            {
                // Row index of the current data row within the worksheet
                int rowIndex = ordersTable.StartRow + 1 + i;
                // Column index of the calculated column within the worksheet
                int colIndex = ordersTable.StartColumn + 2;
                Cell cell = sheet.Cells[rowIndex, colIndex];
                Console.WriteLine($"Cell {cell.Name} formula: {cell.Formula}");
            }

            // Save the workbook (optional, demonstrates that the file contains the calculated column)
            string outputPath = "CalculatedColumn_XLOOKUP.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
