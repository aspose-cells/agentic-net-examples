// Title: Show Zero Values in an Aspose.Cells Pivot Table (C#)
// Description: Creates a workbook, adds sample data with zeros, builds a pivot table, enables Worksheet.DisplayZeros, refreshes and calculates the pivot, and saves the file so zero values appear instead of blanks.
// Keywords: Aspose.Cells | C# | PivotTable | DisplayZeros | show zero values | Excel zero cells | worksheet display zeros | refresh pivot data | calculate pivot data | Aspose.Cells example
// Common Searches: Aspose.Cells show zero values in pivot table | C# display zeros in Excel pivot using Aspose | Worksheet.DisplayZeros property Aspose.Cells | pivot table zero cells Aspose.Cells C# | how to enable zero display in Aspose pivot
// Developer Intent: Configure a pivot table so that cells containing zero are displayed rather than left blank.
// Use Cases: Sales dashboards where categories with zero revenue must be visible. | Financial statements that need to list accounts with a zero balance. | Inventory reports that require items with zero stock to appear in the pivot view.
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table and forces zero values to be shown. | Explain how to use Worksheet.DisplayZeros to display zeros in a pivot table and why RefreshData/CalculateData are required. | Provide step‑by‑step instructions for enabling zero display in an Aspose.Cells pivot table and saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data with zeros, builds a pivot table, enables Worksheet.DisplayZeros, refreshes and calculates the pivot, and saves the file so zero values appear instead of blanks.
    public class PivotTableDisplayZeroValuesDemo
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate source data with some zero values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(0);      // zero value
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(0);      // zero value

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B4", "D2", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Ensure that zero values are displayed in the worksheet
                sheet.DisplayZeros = true;

                // Refresh and calculate the pivot table data using the correct API
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                string outputPath = "PivotTableDisplayZeroValues.xlsx";
                workbook.Save(outputPath);
            }
            catch (Exception ex)
            {
                // Log any runtime errors that occur during Run()
                Console.WriteLine($"Run error: {ex.Message}");
                throw; // Re‑throw to be caught by Main's outer handler if needed
            }
        }
    }
}
