// Title: Aspose.Cells .NET – Display a dash for empty PivotTable cells
// Description: Creates a workbook, adds sample data with blanks, builds a PivotTable, enables custom null‑value display (DisplayNullString = true, NullString = "-"), refreshes the cache, and saves the file as an XLSX document.
// Keywords: Aspose.Cells PivotTable dash for empty cells | DisplayNullString property | NullString Aspose.Cells | C# replace null values in PivotTable | custom null string .NET Excel | Aspose.Cells empty cell placeholder
// Common Searches: Aspose.Cells show dash for null pivot values | C# set PivotTable empty cell text Aspose | DisplayNullString Aspose.Cells example | how to replace blank pivot cells with '-' in .NET | Aspose.Cells custom null string for PivotTable
// Developer Intent: Configure a PivotTable so that any null or empty cell is rendered as a dash (“-”).
// Use Cases: Financial reports where missing amounts must be clearly marked. | Dashboard exports that need a visible placeholder for absent data. | Reusable utility that adds a PivotTable and automatically formats empty values with a custom symbol.
// AI Prompts: Generate C# code using Aspose.Cells to create a PivotTable that displays "-" for empty cells. | Explain how DisplayNullString and NullString affect PivotTable rendering in Aspose.Cells. | Provide a step‑by‑step tutorial for setting a custom null string in an Aspose.Cells PivotTable for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data with blanks, builds a PivotTable, enables custom null‑value display (DisplayNullString = true, NullString = "-"), refreshes the cache, and saves the file as an XLSX document.
    public class PivotTableDisplayDashForEmptyCells
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (including some empty cells)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("");   // Empty category cell
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("C");
                sheet.Cells["B5"].PutValue(null); // Empty value cell

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table fields
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Value as data field

                // Set the pivot table to display a custom string for null/empty cells
                pivotTable.DisplayNullString = true;   // Enable custom null string display
                pivotTable.NullString = "-";           // Use dash character for empty cells

                // Refresh pivot data and calculate results
                pivotTable.RefreshData();   // Correct API to refresh pivot cache
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotTableWithDashForEmptyCells.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotTableDisplayDashForEmptyCells.Run();
        }
    }
}
