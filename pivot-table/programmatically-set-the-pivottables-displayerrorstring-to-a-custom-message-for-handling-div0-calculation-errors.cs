// Title: Custom PivotTable Error Text for #DIV/0! Using Aspose.Cells .NET
// Description: Demonstrates how to replace the default #DIV/0! value in a PivotTable with a custom message. The example creates a workbook, adds a calculated column that can cause division‑by‑zero, builds a PivotTable, enables DisplayErrorString, sets ErrorString to "Division Error", refreshes the pivot and saves the file.
// Keywords: Aspose.Cells PivotTable custom error string | DisplayErrorString property | ErrorString Aspose.Cells | #DIV/0! handling C# | PivotTable error handling .NET | Aspose.Cells calculated column error
// Common Searches: Aspose.Cells replace #DIV/0! in PivotTable | PivotTable DisplayErrorString C# example | Set custom error text for Aspose.Cells pivot | How to hide division errors in Aspose.Cells | Aspose.Cells PivotTable error string usage
// Developer Intent: Show a user‑friendly message instead of #DIV/0! in a PivotTable generated with Aspose.Cells.
// Use Cases: Financial reports where division‑by‑zero should appear as a clear label rather than an Excel error. | Dashboard worksheets that need consistent error wording across multiple PivotTables. | Automated reporting pipelines that must suppress raw error codes for end‑user readability.
// AI Prompts: Write C# code with Aspose.Cells to enable DisplayErrorString and set a custom ErrorString for a PivotTable, then refresh and save the workbook. | Explain the effect of the DisplayErrorString and ErrorString properties on PivotTable error values in Aspose.Cells. | Provide a sample that assigns different custom messages for #DIV/0! and #N/A errors in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotErrorStringDemo
{
    // Demonstrates how to replace the default #DIV/0! value in a PivotTable with a custom message. The example creates a workbook, adds a calculated column that can cause division‑by‑zero, builds a PivotTable, enables DisplayErrorString, sets ErrorString to "Division Error", refreshes the pivot and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate source data
            // Column A – Category, Column B – Value1, Column C – Value2 (will cause division by zero)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value1");
            cells["C1"].PutValue("Value2");

            cells["A2"].PutValue("X");
            cells["B2"].PutValue(10);
            cells["C2"].PutValue(2);   // 10/2 = 5

            cells["A3"].PutValue("Y");
            cells["B3"].PutValue(20);
            cells["C3"].PutValue(0);   // 20/0 -> #DIV/0!

            cells["A4"].PutValue("Z");
            cells["B4"].PutValue(30);
            cells["C4"].PutValue(5);   // 30/5 = 6

            // Add a calculated column that divides Value1 by Value2
            // This will generate a #DIV/0! error for row Y
            cells["D1"].PutValue("Result");
            cells["D2"].Formula = "=B2/C2";
            cells["D3"].Formula = "=B3/C3";
            cells["D4"].Formula = "=B4/C4";

            // Calculate formulas so that error values are materialized
            workbook.CalculateFormula();

            // Create a pivot table based on the source range (including the calculated column)
            int pivotIndex = sheet.PivotTables.Add("A1:D4", "F3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields: Category to rows, Result to data
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Result");

            // Set custom error handling: display a custom string instead of #DIV/0!
            pivot.DisplayErrorString = true;
            pivot.ErrorString = "Division Error";

            // Refresh and calculate pivot data
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableWithCustomErrorString.xlsx");

            Console.WriteLine("PivotTable created with custom error string.");
        }
    }
}
