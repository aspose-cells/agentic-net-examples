// Title: Convert HTML to Excel workbook with active hyperlinks using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells load HTML file and export to XLSX while keeping hyperlinks clickable | How to preserve hyperlink formatting when converting HTML to Excel with Aspose.Cells | Apply hyperlink style to cells after importing HTML into a workbook using Aspose.Cells | Load HTML from a memory stream into an Aspose.Cells workbook and save as .xlsx with active links | Aspose.Cells convert HTML to Excel and retain clickable URLs in C#
// Tags: Aspose.Cells HTML to XLSX conversion | preserve hyperlinks during HTML import | hyperlink styling in Aspose.Cells workbook | load HTML from stream Aspose.Cells | save workbook with active links Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using System.Drawing;

// The example reads an HTML file, loads it into an Aspose.Cells Workbook via a memory stream using LoadFormat.Html, creates a blue underlined style, applies this style to the first cell of each hyperlink range, and saves the workbook as an .xlsx file where all hyperlinks remain clickable.
class HtmlToExcelConverter
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Verify that the HTML file exists
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not found.");
                return;
            }

            // Read the HTML content
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // Load HTML into a memory stream (Aspose.Cells can load from a stream)
            using (MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes(htmlContent)))
            {
                // LoadOptions specify that the source format is HTML
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);

                // Create a Workbook from the HTML stream
                Workbook workbook = new Workbook(htmlStream, loadOptions);

                // Define a typical hyperlink style (blue & underlined)
                Style hyperlinkStyle = workbook.CreateStyle();
                hyperlinkStyle.Font.Underline = FontUnderlineType.Single;
                hyperlinkStyle.Font.Color = Color.Blue;

                // Apply the style to cells that contain hyperlinks
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Hyperlink link in sheet.Hyperlinks)
                    {
                        // Hyperlink.Area provides the cell range; use the first cell of that range
                        int row = link.Area.StartRow;
                        int column = link.Area.StartColumn;

                        Cell cell = sheet.Cells[row, column];
                        cell.SetStyle(hyperlinkStyle);
                    }
                }

                // Save the workbook to an Excel file; hyperlinks remain clickable in the saved file.
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Conversion completed. Excel file saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
