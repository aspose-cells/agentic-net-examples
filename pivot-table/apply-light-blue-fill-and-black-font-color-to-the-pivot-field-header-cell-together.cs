// Title: Aspose.Cells .NET: Apply Light‑Blue Fill and Black Font to a Pivot Table Header Cell
// Description: Creates a workbook, adds a simple pivot table, defines a Style with a solid light‑blue background and black font, and uses PivotTable.Format to style the header cell (row 2, column 0) before saving the file.
// Keywords: Aspose.Cells pivot header style | C# PivotTable.Format | light blue fill Excel | black font pivot header | custom pivot table styling .NET | Aspose.Cells style example
// Common Searches: Aspose.Cells set pivot header background color C# | format pivot table header cell with custom style | apply solid fill to Excel pivot header using Aspose | change font color of pivot field header in .NET | PivotTable.Format example Aspose.Cells
// Developer Intent: Add a light‑blue solid fill and black text to a specific pivot table header cell.
// Use Cases: Highlight row‑field headers in sales dashboards for better visual separation. | Enforce corporate branding by applying a consistent header color scheme across generated reports. | Automate uniform styling of multiple pivot tables in a workbook during batch processing.
// AI Prompts: Generate C# code that creates a Style with a light‑blue fill and black bold font and applies it to a pivot table header cell using Aspose.Cells. | Show how to style all header cells of an Aspose.Cells PivotTable with a single custom Style. | Explain the parameters of PivotTable.Format and how to locate the header row index for styling.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotFieldType

namespace AsposeCellsPivotExample
{
    // Creates a workbook, adds a simple pivot table, defines a Style with a solid light‑blue background and black font, and uses PivotTable.Format to style the header cell (row 2, column 0) before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["A2"].PutValue("Bike");
                worksheet.Cells["B2"].PutValue(5000);
                worksheet.Cells["A3"].PutValue("Car");
                worksheet.Cells["B3"].PutValue(12000);
                worksheet.Cells["A4"].PutValue("Truck");
                worksheet.Cells["B4"].PutValue(8000);

                // Add a pivot table (source range A1:B4, destination top‑left cell D3)
                int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add fields to the pivot table (row field and data field)
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Calculate the pivot data so the layout is generated
                pivotTable.CalculateData();

                // Create a style: light blue fill and black font color
                Style headerStyle = workbook.CreateStyle();
                headerStyle.ForegroundColor = Color.LightBlue;   // fill color
                headerStyle.Pattern = BackgroundType.Solid;      // apply fill
                headerStyle.Font.Color = Color.Black;            // font color
                headerStyle.Font.IsBold = true;                  // optional: make it bold

                // Apply the style to the header cell in the pivot table (row 2, column 0)
                // Row and column indexes are zero‑based.
                pivotTable.Format(2, 0, headerStyle);

                // Determine output path and ensure directory exists
                string outputFile = "PivotHeaderFormatted.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile)) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the pivot table:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
