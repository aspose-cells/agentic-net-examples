// Title: C# Example: Apply Light Blue Fill to a Pivot Table Header with Aspose.Cells
// Description: Creates a workbook, adds sample data, builds a pivot table at D3, defines a solid LightBlue style, and applies it to the pivot table's top‑left header cell using PivotTable.Format before saving the file.
// Keywords: Aspose.Cells | C# | .NET | pivot table header style | background color | light blue fill | PivotTable.Format | Excel export | style object
// Common Searches: Aspose.Cells set pivot table header background color C# | how to format pivot table header cell with Aspose.Cells | apply light blue fill to Excel pivot header using .NET | C# example for styling pivot table header Aspose | PivotTable.Format light blue background Aspose.Cells
// Developer Intent: Add a light blue background fill to a pivot table header cell in a .NET workbook.
// Use Cases: Highlight pivot table headers in automated reports for better visual separation. | Match corporate branding by applying consistent header colors across generated Excel files. | Create reusable code that styles pivot table headers when exporting data programmatically.
// AI Prompts: Generate C# code that uses Aspose.Cells to set a solid light blue fill on a specific pivot table header cell. | Show an Aspose.Cells example that creates a pivot table and formats its header with a custom style via PivotTable.Format. | Explain step‑by‑step how to apply a Style object to a pivot table cell in Aspose.Cells for .NET.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHeaderStyle
{
    // Creates a workbook, adds sample data, builds a pivot table at D3, defines a solid LightBlue style, and applies it to the pivot table's top‑left header cell using PivotTable.Format before saving the file.
    public class ApplyLightBlueHeader
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Food";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["A3"].Value = "Food";
                sheet.Cells["B3"].Value = 80;
                sheet.Cells["A4"].Value = "Drink";
                sheet.Cells["B4"].Value = 150;
                sheet.Cells["A5"].Value = "Drink";
                sheet.Cells["B5"].Value = 70;

                // Add a pivot table based on the data range, placed at D3
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category as row field, Amount as data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column A (Category)
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column B (Amount)

                // Refresh the pivot cache and calculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Header cell of the pivot table is the top‑left cell (D3)
                Cell headerCell = sheet.Cells["D3"];

                // Create a style with light blue fill
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Solid;
                style.ForegroundColor = Color.LightBlue;   // Fill color

                // Apply the style to the header cell via the pivot table
                pivotTable.Format(headerCell.Row, headerCell.Column, style);

                // Ensure output directory exists
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PivotHeaderLightBlue.xlsx");
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the console application
        public static void Main()
        {
            Run();
        }
    }
}
