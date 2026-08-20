// Title: C# – Apply Theme Accent6 to Header Row in a Dynamic Excel Report with Aspose.Cells
// Description: Creates a new workbook, fills header and data rows, defines a style that uses the workbook's Accent6 theme color for both cell fill and font, applies the style to the first row, and saves the file as DynamicReport.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# theme color | Accent6 Excel header style | apply theme color Aspose.Cells | Excel header formatting .NET | dynamic report styling Aspose | ThemeColor Accent6 C# | Excel workbook theme styling
// Common Searches: Aspose.Cells set Accent6 theme color for header | C# apply theme color to Excel row | how to use ThemeColor in Aspose.Cells | format Excel header with theme accent in .NET | apply solid background theme color Aspose.Cells
// Developer Intent: Use the workbook's Accent6 theme color for the background and font of the header row in a generated Excel file.
// Use Cases: Standardize the look of sales or inventory reports by applying the corporate Accent6 color to column titles. | Create multi‑sheet workbooks where each sheet shares a consistent header style based on the workbook theme. | Generate automated Excel exports from applications that need branding without hard‑coding RGB values.
// AI Prompts: Show how to switch the header style to another theme accent (e.g., Accent2) in the same code. | Demonstrate applying the Accent6 style to a specific range like A1:D1 instead of the whole row. | Explain how to combine the Accent6 theme color with a patterned fill or gradient using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ReportGenerator
{
    // Creates a new workbook, fills header and data rows, defines a style that uses the workbook's Accent6 theme color for both cell fill and font, applies the style to the first row, and saves the file as DynamicReport.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data for the report
            string[] headers = { "ID", "Name", "Quantity", "Price" };
            string[,] data = {
                { "1", "Apple",  "50", "0.5" },
                { "2", "Banana", "30", "0.3" },
                { "3", "Cherry", "20", "1.2" }
            };

            // Populate header row (row 0)
            for (int col = 0; col < headers.Length; col++)
            {
                cells[0, col].PutValue(headers[col]);
            }

            // Populate data rows starting from row 1
            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    cells[row + 1, col].PutValue(data[row, col]);
                }
            }

            // Create a style that uses the theme's Accent6 color for background and font
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;
            // Apply Accent6 as the foreground (cell fill) theme color
            headerStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent6, 0);
            // Apply Accent6 as the font theme color
            headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent6, 0);
            headerStyle.Font.IsBold = true;

            // Apply the style to the entire header row
            StyleFlag flag = new StyleFlag { All = true };
            worksheet.Cells.Rows[0].ApplyStyle(headerStyle, flag);

            // Save the workbook
            workbook.Save("DynamicReport.xlsx");
        }
    }
}
