// Title: Insert a Row and Auto‑Recalculate Formulas with Aspose.Cells for .NET (C#)
// Description: This example demonstrates how to create a workbook, add values and formulas, insert a new row at a specific index, and then invoke Workbook.CalculateFormula to refresh all dependent calculations before saving the file as XLSX.
// Keywords: Aspose.Cells | C# | InsertRow | Workbook.CalculateFormula | recalculate Excel formulas | update dependent cells | row insertion automation | SUM formula refresh | XLSX export | sample code
// Common Searches: Aspose.Cells insert row and recalculate formulas | C# Workbook.CalculateFormula usage | How to refresh formulas after adding a row in Aspose.Cells | Update SUM range after row insertion .NET | Excel automation recalculate after structural change
// Developer Intent: Refresh every formula in a workbook after a row is inserted so that all dependent calculations reflect the new data.
// Use Cases: Add a data row between existing entries and automatically adjust SUM or AVERAGE formulas that reference the shifted range. | Programmatically insert rows in a financial report and ensure all linked calculations stay accurate without manual edits. | Perform bulk row insertions in a large spreadsheet and trigger a single recalculation step to maintain performance.
// AI Prompts: Show me C# code that inserts a row at index 1 using Aspose.Cells and then calls Workbook.CalculateFormula. | Explain how Workbook.CalculateFormula updates dependent cells after a row insertion in an Excel workbook. | Provide a step‑by‑step guide to refresh all formulas after inserting multiple rows with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsInsertRowAndRecalculate
{
    // This example demonstrates how to create a workbook, add values and formulas, insert a new row at a specific index, and then invoke Workbook.CalculateFormula to refresh all dependent calculations before saving the file as XLSX.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data and formulas before inserting a row
            // A1 = 10, A2 = 20, B1 = =A1*2, B2 = =A2*2, C1 = =SUM(B1:B2)
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=A1*2";
            cells["B2"].Formula = "=A2*2";
            cells["C1"].Formula = "=SUM(B1:B2)";

            // Insert a new row at index 1 (between the two data rows)
            // This will shift the original row 1 (A2, B2) down to row 2
            cells.InsertRow(1);

            // Optionally put new data into the inserted row
            cells["A1"].PutValue(5); // New value in the inserted row
            cells["B1"].Formula = "=A1*2";

            // Recalculate all formulas so that dependent cells (e.g., C1) reflect the changes
            workbook.CalculateFormula();

            // Save the workbook (save rule)
            workbook.Save("InsertedRowAndRecalculated.xlsx", SaveFormat.Xlsx);
        }
    }
}
