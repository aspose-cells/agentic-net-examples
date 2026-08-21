// Title: Apply a Light Gray Fill Pattern to an Entire Table with Aspose.Cells (C#)
// Description: Demonstrates how to create a workbook, build a 3‑column ListObject, define a style with the Gray25 pattern and LightGray foreground, enable cell shading, and apply the style to every cell of the table—including the header—using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# fill pattern | Excel table background color Aspose | Gray25 cell shading .NET | apply style to ListObject Aspose.Cells | light gray table formatting | Aspose.Cells style flag CellShading | C# Excel table styling
// Common Searches: Aspose.Cells apply gray fill to table | C# set background pattern for ListObject | How to shade all cells in an Aspose.Cells table | Aspose.Cells Gray25 style example | Apply cell shading to Excel table using .NET
// Developer Intent: Add a light gray fill pattern to every cell of an Aspose.Cells table.
// Use Cases: Enhance readability of generated reports by giving the whole data table a subtle gray background. | Maintain consistent visual styling for both header and data rows of a ListObject. | Apply a pattern to a specific range without altering other worksheet formatting.
// AI Prompts: Generate C# code with Aspose.Cells that applies a Gray25 fill pattern and LightGray foreground to a ListObject covering A1:C4. | Show how to change the fill pattern to another BackgroundType while keeping the same range and style flags in Aspose.Cells. | Explain how to apply the light gray fill pattern only to the data rows of a table, excluding the header, using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableFillPattern
{
    // Demonstrates how to create a workbook, build a 3‑column ListObject, define a style with the Gray25 pattern and LightGray foreground, enable cell shading, and apply the style to every cell of the table—including the header—using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (A1:C4)
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Category");
                worksheet.Cells["C1"].PutValue("Price");

                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue("Fruit");
                worksheet.Cells["C2"].PutValue(1.20);

                worksheet.Cells["A3"].PutValue("Carrot");
                worksheet.Cells["B3"].PutValue("Vegetable");
                worksheet.Cells["C3"].PutValue(0.80);

                worksheet.Cells["A4"].PutValue("Bread");
                worksheet.Cells["B4"].PutValue("Bakery");
                worksheet.Cells["C4"].PutValue(2.50);

                // Create a ListObject (Excel table) covering the data range
                // Parameters: startRow, startColumn, totalRows, totalColumns, hasHeaders
                int listObjectIndex = worksheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = worksheet.ListObjects[listObjectIndex];
                table.ShowHeaderRow = true;
                table.ShowTableStyleFirstColumn = true;
                table.ShowTableStyleRowStripes = true;

                // Define a style with a light gray fill pattern
                Style style = workbook.CreateStyle();
                style.Pattern = BackgroundType.Gray25;               // Light gray pattern
                style.ForegroundColor = Color.LightGray;            // Foreground color for the pattern
                style.BackgroundColor = Color.White;                // Background color (optional)

                // Enable cell shading in the style flag
                StyleFlag flag = new StyleFlag();
                flag.CellShading = true;

                // Apply the style to the entire table range (including header)
                // The table occupies rows 0‑4 (5 rows) and columns 0‑2 (3 columns)
                Aspose.Cells.Range tableRange = worksheet.Cells.CreateRange(0, 0, 5, 3);
                tableRange.ApplyStyle(style, flag);

                // Save the workbook
                string outputPath = "TableLightGrayPattern.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
