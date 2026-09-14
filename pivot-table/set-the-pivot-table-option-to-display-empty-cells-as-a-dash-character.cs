// Title: Configure Aspose.Cells PivotTable in C# to display a dash for empty cells
// AI Prompts: Generate a workbook, populate it with sample data containing null entries, add a pivot table, enable DisplayNullString, assign '-' to NullString, refresh and calculate the pivot, then save the workbook. | Configure a PivotTable in Aspose.Cells C# to show '-' for empty values by setting DisplayNullString = true and NullString = '-'.
// Common Searches: Aspose.Cells C# pivot table replace blank cells with dash | Set custom null string for pivot table values using Aspose.Cells API | Display '-' for empty data in Aspose.Cells pivot table C# example | How to show placeholder for null cells in Aspose.Cells pivot table
// Tags: Aspose.Cells PivotTable DisplayNullString usage | C# set NullString dash for empty pivot cells | Aspose.Cells custom empty cell placeholder in pivot tables | PivotTable empty value formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data with null entries, inserts a pivot table, enables DisplayNullString, sets NullString to '-', refreshes and calculates the pivot, and saves the result as an .xlsx file.
    public class PivotTableDisplayEmptyAsDash
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (including some null/empty cells)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["A4"].PutValue(null); // Empty category cell
                worksheet.Cells["B4"].PutValue(null); // Empty value cell

                // Add a pivot table based on the data range
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

                // Set the pivot table to display a custom string for empty (null) cells
                pivotTable.DisplayNullString = true; // Enable custom null string display
                pivotTable.NullString = "-";         // Use dash character for empty cells

                // Refresh the pivot cache and calculate the pivot table to apply changes
                pivotTable.RefreshData();   // Correct API to refresh data source
                pivotTable.CalculateData();

                // Save the workbook with the configured pivot table
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotTableEmptyAsDash.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred in Run(): {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                PivotTableDisplayEmptyAsDash.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
