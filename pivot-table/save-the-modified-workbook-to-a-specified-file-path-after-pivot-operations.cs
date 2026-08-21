// Title: C# – Save an Aspose.Cells Workbook with a Pivot Table to a Specified Path
// Description: Demonstrates how to create a workbook, populate it with sample data, add and configure a PivotTable, disable pivot cache storage, and persist the file to a custom location using Workbook.Save(string) in Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# save workbook | pivot table export Aspose | Workbook.Save custom path | disable pivot cache Aspose.Cells | C# Excel pivot table example | Aspose.Cells save to .xlsx | export pivot data programmatically
// Common Searches: How to save an Aspose.Cells workbook that contains a pivot table in C# | Disable pivot cache when saving Excel with Aspose.Cells | Specify output file path for Workbook.Save after creating a pivot | Aspose.Cells C# save pivot table to file | Workbook.Save(string) example with PivotTable
// Developer Intent: Persist a workbook that includes a configured PivotTable to a user‑defined file location while optionally omitting the pivot cache.
// Use Cases: Generate a sales‑summary PivotTable and export it as an .xlsx file for reporting pipelines. | Create a temporary workbook, add a PivotTable, turn off cache to reduce file size, and save it to a user‑chosen directory. | Automate dashboard generation where each run produces a uniquely named Excel file containing pivot analysis.
// AI Prompts: Write C# code that builds a workbook, adds a PivotTable, sets SaveData = false, and saves the file to a given path using Aspose.Cells. | Explain how to change the output format (e.g., .xls, .csv) when calling Workbook.Save on a workbook that contains a PivotTable. | Show how to programmatically set the file path and name for a saved Aspose.Cells workbook with a pivot cache disabled.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotSaveDemo
{
    // Demonstrates how to create a workbook, populate it with sample data, add and configure a PivotTable, disable pivot cache storage, and persist the file to a custom location using Workbook.Save(string) in Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Beverage");
            sheet.Cells["B3"].PutValue(800);
            sheet.Cells["A4"].PutValue("Electronics");
            sheet.Cells["B4"].PutValue(1500);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: rows = Category, data = Sum of Amount
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh the pivot table to calculate its data
            sheet.RefreshPivotTables();

            // Optionally, disable saving pivot cache data with the workbook
            pivot.SaveData = false;

            // Save the modified workbook to the desired file path
            string outputPath = "ModifiedPivotWorkbook.xlsx";
            workbook.Save(outputPath); // Uses Workbook.Save(string) rule

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
    }
}
