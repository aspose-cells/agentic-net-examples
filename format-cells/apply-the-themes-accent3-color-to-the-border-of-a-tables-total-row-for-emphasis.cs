// Title: Apply Theme Accent3 Border to a Table Total Row with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a ListObject with a totals row, builds a custom table style, and applies a thick Accent3 theme border to all sides of the total row before saving the file.
// Keywords: Aspose.Cells | C# | theme Accent3 | total row border | custom table style | ListObject | Excel border color | thick border | table style customization | Excel automation
// Common Searches: Aspose.Cells set total row border color | how to apply theme accent to table total row in C# | custom table style with thick borders Aspose.Cells | add border to ListObject total row using theme color | Aspose.Cells table style Accent3 example
// Developer Intent: Create a reusable table style that highlights the totals row with a thick Accent3 theme border.
// Use Cases: Emphasize financial summary rows in generated Excel reports. | Standardize total‑row appearance across multiple workbooks in an enterprise solution. | Produce inventory sheets where the summed quantity row stands out visually.
// AI Prompts: Show how to change the border to Accent2 instead of Accent3. | Demonstrate applying a dashed, theme‑colored border to the header row of a table. | Explain how to reuse the Accent3 total‑row style for several tables on the same worksheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a ListObject with a totals row, builds a custom table style, and applies a thick Accent3 theme border to all sides of the total row before saving the file.
    public class TableTotalRowAccent3Border
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (3 data rows + header)
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Quantity");
                cells["A2"].PutValue("Apple");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("Banana");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("Cherry");
                cells["B4"].PutValue(15);

                // Add a table that includes the data range; the last row will become the total row
                int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTotals = true; // Enable total row
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum; // Sum on Quantity column

                // ------------------------------------------------------------
                // Create a custom table style to emphasize the total row border
                // ------------------------------------------------------------

                // Create a style for the total row
                Style totalRowStyle = workbook.CreateStyle();

                // Create a CellsColor with the theme Accent3 color (no tint)
                CellsColor accent3Color = workbook.CreateCellsColor();
                accent3Color.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

                // Apply the theme color to all four borders of the style
                totalRowStyle.SetBorder(BorderType.TopBorder, CellBorderType.Thick, accent3Color);
                totalRowStyle.SetBorder(BorderType.BottomBorder, CellBorderType.Thick, accent3Color);
                totalRowStyle.SetBorder(BorderType.LeftBorder, CellBorderType.Thick, accent3Color);
                totalRowStyle.SetBorder(BorderType.RightBorder, CellBorderType.Thick, accent3Color);

                // Add a new custom table style to the workbook's TableStyles collection
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                string customStyleName = "Accent3TotalRowStyle";
                int styleIdx = tableStyles.AddTableStyle(customStyleName);
                TableStyle customTableStyle = tableStyles[styleIdx];

                // Add the TotalRow element to the custom style and assign the created style
                TableStyleElementCollection elements = customTableStyle.TableStyleElements;
                int elementIdx = elements.Add(TableStyleElementType.TotalRow);
                TableStyleElement totalRowElement = elements[elementIdx];
                totalRowElement.SetElementStyle(totalRowStyle);

                // Apply the custom style to the table
                table.TableStyleName = customStyleName;

                // Save the workbook
                string outputPath = "TableTotalRowAccent3Border.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableTotalRowAccent3Border.Run();
        }
    }
}
