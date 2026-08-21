// Title: Aspose.Cells .NET: Retrieve Slicer Properties, Log Them, and Increase Row Height
// Description: Creates a workbook with a pivot table, adds a slicer for the "Category" field, prints slicer attributes (Name, RowHeight, RowHeightPixel, NumberOfColumns, Caption) to the console, then raises the slicer's RowHeight by 5 points before saving the file.
// Keywords: Aspose.Cells slicer properties | C# get slicer row height | modify slicer dimensions .NET | log slicer attributes Aspose | pivot table slicer example | adjust slicer RowHeight programmatically
// Common Searches: how to read slicer properties with Aspose.Cells | increase slicer row height in C# | Aspose.Cells log slicer details to console | change slicer size after creation .NET | retrieve slicer caption and column count Aspose
// Developer Intent: Read slicer attributes, output them for inspection, and programmatically enlarge the slicer's row height.
// Use Cases: Audit slicer layout by extracting current settings from a workbook. | Dynamically resize slicers to improve readability based on their existing dimensions. | Debug pivot‑table interactions by logging slicer metadata during development.
// AI Prompts: Write C# code that extracts all slicer properties from an Aspose.Cells workbook and saves the data to a CSV file. | Show how to set slicer RowHeight proportionally to the number of items it contains using Aspose.Cells for .NET. | Provide best‑practice guidelines for handling exceptions when accessing slicer properties in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Creates a workbook with a pivot table, adds a slicer for the "Category" field, prints slicer attributes (Name, RowHeight, RowHeightPixel, NumberOfColumns, Caption) to the console, then raises the slicer's RowHeight by 5 points before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("Fruit");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("Fruit");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("Vegetable");
            worksheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount column

            // Add a slicer linked to the pivot table for the "Category" field.
            // Destination cell must be a valid cell address (e.g., "F1").
            int slicerIndex = worksheet.Slicers.Add(pivotTable, "F1", "Category");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Set an initial row height for the slicer (points)
            slicer.RowHeight = 15;

            // Retrieve and log slicer properties
            Console.WriteLine($"Slicer Name: {slicer.Name}");
            Console.WriteLine($"RowHeight (points): {slicer.RowHeight}");
            Console.WriteLine($"RowHeightPixel: {slicer.RowHeightPixel}");
            Console.WriteLine($"NumberOfColumns: {slicer.NumberOfColumns}");
            Console.WriteLine($"Caption: {slicer.Caption}");

            // Modify the row height based on the logged value (increase by 5 points)
            double updatedHeight = slicer.RowHeight + 5;
            slicer.RowHeight = updatedHeight;
            Console.WriteLine($"Updated RowHeight (points): {slicer.RowHeight}");

            // Save the workbook
            workbook.Save("SlicerPropertiesDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
