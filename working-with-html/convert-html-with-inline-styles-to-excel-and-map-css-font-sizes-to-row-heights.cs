// Title: Convert HTML with inline CSS font-size styles to Excel and set row heights based on font size using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an HTML string containing inline font-size styles into an Aspose.Cells Workbook and automatically adjusts each row's height to match the largest font size in that row. | Show how to calculate Excel row height from CSS point values with a configurable multiplier when importing HTML via Aspose.Cells. | Write a reusable C# method that accepts HTML content, imports it into a Workbook, and returns the workbook with row heights calibrated to the CSS font-size values.
// Common Searches: asp.net convert html table with inline font-size to excel row height using Aspose.Cells | c# map css point font-size to excel row height aspose.cells | load html from memory stream and preserve font size in excel rows aspose.cells .net | adjust row heights after importing html with inline styles aspose.cells
// Tags: html to excel conversion aspose.cells | inline css font-size mapping to row height | row height calculation from font size c# | memory stream html load aspose.cells | configurable row height multiplier aspose.cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

// // Loads an HTML string with inline font-size styles into an Aspose.Cells Workbook, determines the maximum font size per row, applies a multiplier to set each row's height accordingly, and saves the result as an XLSX file.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Sample HTML with inline CSS font-size styles
            string html = @"
            <html>
            <body>
            <table>
                <tr><td style='font-size:12pt;'>Small</td></tr>
                <tr><td style='font-size:16pt;'>Medium</td></tr>
                <tr><td style='font-size:20pt;'>Large</td></tr>
            </table>
            </body>
            </html>";

            // Load the HTML content into a new workbook using a memory stream
            Workbook workbook;
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(html)))
            {
                HtmlLoadOptions loadOptions = new HtmlLoadOptions();
                workbook = new Workbook(ms, loadOptions);
            }

            // Reference to the first worksheet where HTML was imported
            Worksheet sheet = workbook.Worksheets[0];

            // Store the maximum font size found in each row
            Dictionary<int, double> rowMaxFontSize = new Dictionary<int, double>();

            // Scan all cells to capture font sizes from inline styles
            foreach (Cell cell in sheet.Cells)
            {
                double fontSize = cell.GetStyle().Font.Size;
                if (fontSize > 0)
                {
                    int rowIndex = cell.Row;
                    if (!rowMaxFontSize.ContainsKey(rowIndex) || fontSize > rowMaxFontSize[rowIndex])
                    {
                        rowMaxFontSize[rowIndex] = fontSize;
                    }
                }
            }

            // Convert font size to row height using a multiplier for visual spacing
            const double multiplier = 1.2;
            foreach (var kvp in rowMaxFontSize)
            {
                int rowIndex = kvp.Key;
                double maxFontSize = kvp.Value;
                sheet.Cells.Rows[rowIndex].Height = maxFontSize * multiplier;
            }

            // Save the workbook to an Excel file
            string outputPath = "Output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
