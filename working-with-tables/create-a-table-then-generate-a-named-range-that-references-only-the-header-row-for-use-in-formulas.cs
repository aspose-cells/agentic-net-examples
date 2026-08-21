// Title: Aspose.Cells for .NET – Create an Excel Table and a Header‑Only Named Range (C#)
// Description: C# example that builds a ListObject table, computes its column count, defines a named range covering only the header row, uses it in a COUNTA formula, and saves the workbook.
// Keywords: Aspose.Cells C# table header named range | ListObject header only range .NET | Excel named range for table headers | create named range Aspose.Cells | C# Aspose.Cells table example | header row range formula
// Common Searches: how to name only the header row of a table in Aspose.Cells | Aspose.Cells create named range for ListObject header | C# Aspose.Cells table column count without ColumnCount property | use table header range in Excel formula with Aspose
// Developer Intent: Generate a table and a named range that points exclusively to its header row for formula references.
// Use Cases: Count or validate column titles with COUNTA, MATCH, or VLOOKUP using the header‑only range. | Populate data‑validation dropdowns with table column names extracted from the named range. | Dynamically read header values for report generation or UI controls without hard‑coding column names.
// AI Prompts: Provide C# code using Aspose.Cells that adds a ListObject, calculates the number of columns, creates a named range for the header row, and applies it in a formula. | Explain why the column count must be derived manually when creating a header‑only named range in Aspose.Cells. | Show how to use the header‑only named range in a data‑validation list or a lookup function such as VLOOKUP.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables; // For ListObject
using AsposeRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

// C# example that builds a ListObject table, computes its column count, defines a named range covering only the header row, uses it in a COUNTA formula, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate header row and some data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Price");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(2.5);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(1.8);

            // Add a table that includes the header row (A1:B3)
            int tableIdx = sheet.ListObjects.Add("A1", "B3", true);
            ListObject table = sheet.ListObjects[tableIdx];
            table.DisplayName = "ProductTable";

            // Calculate column count manually (ColumnCount property not available)
            int columnCount = table.EndColumn - table.StartColumn + 1;

            // Create a named range that references only the header row of the table
            AsposeRange headerRange = cells.CreateRange(
                table.StartRow,          // First row of the table (header)
                table.StartColumn,       // First column of the table
                1,                       // Only one row (the header)
                columnCount);            // Number of columns in the table
            headerRange.Name = "ProductHeaders";

            // Example usage of the named header range in a formula
            cells["C1"].Formula = "=COUNTA(ProductHeaders)";
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save("TableWithHeaderNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
