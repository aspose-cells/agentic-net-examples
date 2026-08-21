// Title: Hide Empty Rows in an Aspose.Cells PivotTable – C# Example
// Description: Shows how to build a workbook, insert sample data with blank rows, create a PivotTable, and suppress those empty rows by setting the ShowEmptyRow property to false with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | PivotTable | Hide empty rows | ShowEmptyRow | C# | .NET | blank rows | Excel automation | Aspose.Cells API | pivot table formatting
// Common Searches: Aspose.Cells hide empty rows in pivot table C# | ShowEmptyRow property Aspose.Cells example | remove blank rows from PivotTable using Aspose.Cells | C# code to hide empty rows in Excel pivot | Aspose.Cells PivotTable empty row handling
// Developer Intent: Configure a PivotTable so that rows without data are not displayed.
// Use Cases: Generate a sales summary where categories with no sales are omitted. | Create a financial report that excludes blank account entries for a cleaner layout. | Build an inventory analysis pivot that automatically removes empty product rows.
// AI Prompts: Write C# code with Aspose.Cells that creates a PivotTable and hides empty rows by adjusting the appropriate property. | Explain the effect of the ShowEmptyRow property in Aspose.Cells and how it differs from other pivot‑table display options. | Provide a step‑by‑step tutorial for creating a workbook, adding data with blank rows, building a PivotTable, and suppressing empty rows before saving the file.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, insert sample data with blank rows, create a PivotTable, and suppress those empty rows by setting the ShowEmptyRow property to false with Aspose.Cells for .NET.
    public class HideEmptyRowsInPivotTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];

                // Populate sample data (including empty rows)
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Product");
                dataSheet.Cells["C1"].PutValue("Sales");

                dataSheet.Cells["A2"].PutValue("Electronics");
                dataSheet.Cells["B2"].PutValue("TV");
                dataSheet.Cells["C2"].PutValue(1000);

                // Empty row (should be hidden in the pivot)
                dataSheet.Cells["A3"].PutValue("");
                dataSheet.Cells["B3"].PutValue("");
                dataSheet.Cells["C3"].PutValue("");

                dataSheet.Cells["A4"].PutValue("Furniture");
                dataSheet.Cells["B4"].PutValue("Chair");
                dataSheet.Cells["C4"].PutValue(500);

                // Add a second empty row
                dataSheet.Cells["A5"].PutValue("");
                dataSheet.Cells["B5"].PutValue("");
                dataSheet.Cells["C5"].PutValue("");

                // Create a new worksheet for the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

                // Add the pivot table (source range A1:C5, destination start cell E3)
                int pivotIndex = pivotSheet.PivotTables.Add("=Sheet1!A1:C5", "E3", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Hide empty rows in the pivot table
                pivotTable.ShowEmptyRow = false;

                // Calculate the pivot data
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTable_HideEmptyRows.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point required for compilation
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
