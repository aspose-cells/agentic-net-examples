// Title: Convert HTML Table to Excel (XLSX) with Locale‑Aware Date Formatting Using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves an HTML string to a temporary file, loads it into an Aspose.Cells Workbook with HtmlLoadOptions specifying a CultureInfo (e.g., en‑GB), and exports the workbook as an XLSX file. | Show how to loop through every cell after importing HTML and assign a custom number format that matches the CultureInfo's short date pattern to any DateTime values. | Demonstrate robust error handling and cleanup of the temporary HTML file when converting HTML to Excel with Aspose.Cells.
// Common Searches: how to import an HTML table into Excel using Aspose.Cells with British date format | c# Aspose.Cells HtmlLoadOptions CultureInfo example for date parsing | convert HTML to xlsx while preserving locale-specific date formatting in .NET | detect DateTime cells after loading HTML with Aspose.Cells and apply custom date style
// Tags: html-to-xlsx conversion Aspose.Cells | locale-aware date parsing Aspose.Cells | HtmlLoadOptions CultureInfo usage | apply custom date format to cells | temporary file handling C# Aspose.Cells

using System;
using System.Globalization;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample program writes an HTML string containing a table to a temporary file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions with the en‑GB culture, scans all cells to find DateTime values, applies the culture's short date pattern as a custom style, saves the workbook as ConvertedFromHtml.xlsx, and finally removes the temporary HTML file.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Sample HTML content (replace with actual HTML source)
        string htmlContent = @"
            <html>
                <body>
                    <table>
                        <tr><td>01/02/2023</td><td>Value 1</td></tr>
                        <tr><td>03/04/2023</td><td>Value 2</td></tr>
                    </table>
                </body>
            </html>";

        // Temporary HTML file path
        string tempHtmlPath = Path.Combine(Path.GetTempPath(), "tempHtml.html");

        try
        {
            // Write HTML content to a temporary file
            File.WriteAllText(tempHtmlPath, htmlContent, Encoding.UTF8);

            // Configure HTML load options with culture info for date parsing
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                Encoding = Encoding.UTF8,
                // Specify the culture used for parsing dates/numbers
                CultureInfo = new CultureInfo("en-GB")
            };

            // Load the HTML file into a new workbook
            Workbook workbook = null;
            if (File.Exists(tempHtmlPath))
            {
                // Use constructor that accepts file path and load options
                workbook = new Workbook(tempHtmlPath, loadOptions);
            }
            else
            {
                Console.WriteLine("Temporary HTML file not found.");
                return;
            }

            // Apply proper date format based on the specified culture
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                foreach (Cell cell in cells)
                {
                    if (cell.Value is DateTime)
                    {
                        Style style = cell.GetStyle();
                        // Use the short date pattern of the load options culture
                        string localeDatePattern = loadOptions.CultureInfo.DateTimeFormat.ShortDatePattern;
                        style.Custom = localeDatePattern;
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook to an Excel file
            workbook.Save("ConvertedFromHtml.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up temporary file
            if (File.Exists(tempHtmlPath))
            {
                try { File.Delete(tempHtmlPath); } catch { /* ignore cleanup errors */ }
            }
        }
    }
}
