// Title: C# – Apply Accent 3 Theme Border to a Table Total Row using Aspose.Cells
// Description: Demonstrates how to build a workbook, insert a ListObject table with a totals row, create a custom TableStyle called Accent3TotalRowStyle, set a thick Accent 3 theme border on all sides of the TotalRow element, apply the style to the table, and save the result as TableTotalRowAccent3Border.xlsx.
// Keywords: Aspose.Cells | C# | .NET | TableStyle | TotalRow | Accent3 | theme border | thick border | custom table style | ListObject | Excel export | border color | theme color | TableTotalRowAccent3Border
// Common Searches: Aspose.Cells set Accent3 border on total row | C# custom TableStyle for TotalRow element | how to add thick theme borders to Excel table totals row | apply theme accent color to table border Aspose.Cells | create reusable table style for total row in .NET
// Developer Intent: Create a custom TableStyle that adds a thick Accent 3 themed border to a table’s total row and apply it to a ListObject.
// Use Cases: Emphasize summary rows in financial reports with a bold Accent 3 border for quick visual identification. | Define a reusable style that automatically formats total rows across multiple worksheets in an automated reporting solution. | Generate Excel files where the totals row stands out using theme‑based colors, ensuring consistency with the workbook’s overall design.
// AI Prompts: Write C# code with Aspose.Cells to give the TotalRow of a ListObject a thick Accent 3 border. | Show how to create and reuse a custom TableStyle named Accent3TotalRowStyle that formats the TotalRow element with a theme accent color. | Provide an example that adds a totals row to an Excel table and styles its borders using the Accent 3 theme in .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Demonstrates how to build a workbook, insert a ListObject table with a totals row, create a custom TableStyle called Accent3TotalRowStyle, set a thick Accent 3 theme border on all sides of the TotalRow element, apply the style to the table, and save the result as TableTotalRowAccent3Border.xlsx.
    public class TableTotalRowAccent3Border
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (2 columns, 3 data rows)
                worksheet.Cells["A1"].PutValue("Product");
                worksheet.Cells["B1"].PutValue("Price");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B3"].PutValue(15);
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["B4"].PutValue(20);

                // Add a table that includes the data range and enable totals row
                int tableIndex = worksheet.ListObjects.Add(0, 0, 3, 1, true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.ShowTotals = true;
                // Set totals calculation for each column
                table.ListColumns[0].TotalsCalculation = TotalsCalculation.Count;
                table.ListColumns[1].TotalsCalculation = TotalsCalculation.Sum;

                // ------------------------------------------------------------
                // Create a custom table style to emphasize the total row border
                // ------------------------------------------------------------
                string customStyleName = "Accent3TotalRowStyle";

                // Access the collection of table styles and add a new one
                TableStyleCollection tableStyles = workbook.Worksheets.TableStyles;
                int styleIdx = tableStyles.AddTableStyle(customStyleName);
                TableStyle customStyle = tableStyles[styleIdx];

                // Add a style element for the TotalRow
                TableStyleElementCollection elements = customStyle.TableStyleElements;
                int elementIdx = elements.Add(TableStyleElementType.TotalRow);
                TableStyleElement totalRowElement = elements[elementIdx];

                // Create a style that will be applied to the total row
                Style totalRowStyle = workbook.CreateStyle();

                // Define the border style (Thick) and set the theme color to Accent3
                // Apply to all four borders
                Border topBorder = totalRowStyle.Borders[BorderType.TopBorder];
                topBorder.LineStyle = CellBorderType.Thick;
                topBorder.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

                Border bottomBorder = totalRowStyle.Borders[BorderType.BottomBorder];
                bottomBorder.LineStyle = CellBorderType.Thick;
                bottomBorder.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

                Border leftBorder = totalRowStyle.Borders[BorderType.LeftBorder];
                leftBorder.LineStyle = CellBorderType.Thick;
                leftBorder.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

                Border rightBorder = totalRowStyle.Borders[BorderType.RightBorder];
                rightBorder.LineStyle = CellBorderType.Thick;
                rightBorder.ThemeColor = new ThemeColor(ThemeColorType.Accent3, 0);

                // Assign the prepared style to the TotalRow element
                totalRowElement.SetElementStyle(totalRowStyle);

                // Apply the custom style to the table
                table.TableStyleName = customStyleName;

                // Save the workbook
                workbook.Save("TableTotalRowAccent3Border.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            TableTotalRowAccent3Border.Run();
        }
    }
}
