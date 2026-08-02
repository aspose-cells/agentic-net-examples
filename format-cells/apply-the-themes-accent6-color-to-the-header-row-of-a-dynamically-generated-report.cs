// Title: C# – Apply Theme Accent6 Color to Header Row in Excel with Aspose.Cells
// Description: Creates a workbook, writes header cells, adds sample data, defines a style that uses the workbook’s Accent6 theme color for both background and font, applies the style to the first row, and saves the file as ReportWithAccent6Header.xlsx.
// Keywords: Aspose.Cells C# theme color | Accent6 header style | Excel header background theme | Apply theme color Aspose.Cells | StyleFlag row styling | C# Excel report formatting | ThemeColorType Accent6 | dynamic Excel report Aspose | set header font theme color | solid background theme color
// Common Searches: Aspose.Cells set header row background to Accent6 | C# apply theme color to Excel row | How to use ThemeColor Accent6 in Aspose.Cells | Apply style to entire row Aspose.Cells C# | Excel report header theme color Aspose.Cells | Create styled header with theme colors in .NET
// Developer Intent: Apply the workbook’s Accent6 theme color to the header row of a generated Excel report.
// Use Cases: Generate a sales summary where the header row follows the corporate Accent6 branding. | Export inventory data from a web service with a themed header that matches the document’s default palette. | Create automated Excel reports in a Windows service that automatically adopt the workbook’s Accent6 color for consistency. | Build a multi‑sheet financial workbook where each sheet’s header uses the same Accent6 style for unified appearance.
// AI Prompts: Show how to add a thin border to the Accent6‑styled header row while keeping the background and font colors unchanged. | Generate C# code that creates a custom theme, sets Accent6 as the header background, and applies the style to several worksheets in one workbook. | Explain how to read the current Accent6 color value from the workbook theme and use it in a conditional formatting rule with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // Creates a workbook, writes header cells, adds sample data, defines a style that uses the workbook’s Accent6 theme color for both background and font, applies the style to the first row, and saves the file as ReportWithAccent6Header.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate header row
                string[] headers = { "ID", "Name", "Quantity", "Price", "Total" };
                for (int col = 0; col < headers.Length; col++)
                {
                    cells[0, col].PutValue(headers[col]);
                }

                // Populate sample data rows
                for (int row = 1; row <= 5; row++)
                {
                    cells[row, 0].PutValue(row);                     // ID
                    cells[row, 1].PutValue($"Item {row}");           // Name
                    cells[row, 2].PutValue(row * 2);                 // Quantity
                    cells[row, 3].PutValue(row * 10.5);              // Price
                    // Total = Quantity * Price
                    cells[row, 4].Formula = $"C{row + 1}*D{row + 1}";
                }

                // Create a style that uses the theme's Accent6 color for background and font
                Style headerStyle = workbook.CreateStyle();
                headerStyle.Pattern = BackgroundType.Solid;
                headerStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent6, 0);
                headerStyle.Font.IsBold = true;
                headerStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent6, 0);

                // Apply the style to the entire header row
                Row headerRow = worksheet.Cells.Rows[0];
                StyleFlag flag = new StyleFlag { All = true };
                headerRow.ApplyStyle(headerStyle, flag);

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportWithAccent6Header.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
