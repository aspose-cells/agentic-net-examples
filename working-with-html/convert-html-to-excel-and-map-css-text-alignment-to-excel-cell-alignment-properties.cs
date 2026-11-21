// Title: Convert an HTML file to XLSX and transfer CSS text‑align values to Excel cell alignment using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load HTML file while preserving CSS text-align in Excel | how to map CSS text-align property to Aspose.Cells cell horizontal alignment | convert HTML table to Excel workbook keeping original alignment with Aspose.Cells | Aspose.Cells LoadOptions Html format example for alignment retention in .NET | sample code to adjust cell style after importing HTML into a workbook
// Tags: Aspose.Cells HTML to XLSX conversion with alignment mapping | LoadOptions Html format usage in Aspose.Cells | cell style horizontal alignment mapping from CSS | C# preserve CSS text-align during HTML import | adjust cell style after HTML load Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an HTML file into an Aspose.Cells Workbook using LoadOptions set to Html format, iterates through the cells to inspect or modify their Style for horizontal alignment, ensures the output directory exists, and saves the result as an XLSX file, demonstrating how to retain or adjust CSS text‑align values in the resulting Excel worksheet.
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

            // Load the HTML file into a workbook using LoadOptions for HTML format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // OPTIONAL: Inspect or adjust cell alignments if needed
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            foreach (Cell cell in cells)
            {
                Style style = cell.GetStyle();

                // Example placeholder for custom alignment adjustments
                // if (style.HorizontalAlignment == TextAlignmentType.General)
                // {
                //     // style.HorizontalAlignment = TextAlignmentType.Center;
                //     // cell.SetStyle(style);
                // }
            }

            // Ensure the output directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to an Excel file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML has been converted to Excel and saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
