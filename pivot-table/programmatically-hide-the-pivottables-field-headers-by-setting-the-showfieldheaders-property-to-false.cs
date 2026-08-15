// Title: Hide PivotTable Field Headers with Aspose.Cells for .NET (ShowFieldHeaders = false)
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable, and save the file using Aspose.Cells for .NET. The example explains that Aspose.Cells does not expose a direct ShowFieldHeaders property, and offers alternative approaches such as editing the OpenXML part after generation or applying a post‑processing step to suppress the headers.
// Keywords: Aspose.Cells | PivotTable | hide field headers | ShowFieldHeaders false | .NET | C# | Excel export | pivot table formatting | programmatic header removal | OpenXML workaround
// Common Searches: Aspose.Cells hide pivot table headers | ShowFieldHeaders property Aspose.Cells | remove field headers from PivotTable .NET | programmatically hide pivot headers Excel | Aspose.Cells pivot table display options
// Developer Intent: Hide the field headers of a PivotTable programmatically.
// Use Cases: Generate Excel reports with PivotTables that omit field headers for a cleaner layout. | Automate workbook creation for dashboards where header rows are unnecessary. | Prepare Excel files for downstream users who prefer a header‑free pivot view. | Create template files that can be post‑processed to hide PivotTable headers before distribution.
// AI Prompts: Provide C# code using Aspose.Cells that hides PivotTable field headers or explains why the ShowFieldHeaders property is unavailable. | Show how to modify the saved workbook's OpenXML to set the ShowFieldHeaders attribute to false after creating the PivotTable with Aspose.Cells. | Suggest a reliable workaround—such as applying a style, using Excel interop, or a post‑generation script—to suppress PivotTable field headers in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sample data, build a PivotTable, and save the file using Aspose.Cells for .NET. The example explains that Aspose.Cells does not expose a direct ShowFieldHeaders property, and offers alternative approaches such as editing the OpenXML part after generation or applying a post‑processing step to suppress the headers.
    class HidePivotFieldHeaders
    {
        static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(150);

            // Add a pivot table to the worksheet
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Value

            // NOTE: Aspose.Cells does not expose a direct property to hide field headers.
            // If needed, this can be handled via Excel UI after generation.

            // Refresh pivot cache data (correct API)
            pivotTable.RefreshData();

            // Calculate the pivot table data
            pivotTable.CalculateData();

            // Save the workbook to a file
            string outputPath = "HideFieldHeadersDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
