// Title: Convert an HTML file to an XLSX workbook with clickable hyperlinks using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a local HTML document into an Aspose.Cells Workbook, keeps all anchor tags as active hyperlinks, and saves the workbook as an .xlsx file. | Write a C# snippet that walks through each worksheet's Hyperlinks collection after importing HTML and prints the cell address together with its URL. | Provide C# logic to verify the target output folder exists (or create it) before invoking Workbook.Save with SaveFormat.Xlsx.
// Common Searches: Aspose.Cells C# convert HTML to Excel keep hyperlinks clickable | how to import HTML with anchor tags into an Aspose.Cells workbook | list hyperlinks after loading HTML into Aspose.Cells worksheet | save workbook as XLSX after HTML conversion with active links using Aspose.Cells | C# create output directory before saving Excel file with Aspose.Cells
// Tags: HTML to XLSX conversion Aspose.Cells | preserve hyperlinks LoadOptions Html | iterate worksheet Hyperlinks collection C# | ensure output folder before Workbook.Save | exception handling Aspose.Cells HTML import

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source HTML file, loads it into an Aspose.Cells Workbook using LoadOptions with Html format, optionally logs each imported hyperlink's cell location and URL, creates the output directory if it does not exist, and saves the workbook as an XLSX file while preserving clickable links, with basic error handling.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlFilePath = "input.html";
        string excelFilePath = "output.xlsx";

        try
        {
            // Verify that the source HTML file exists
            if (!File.Exists(htmlFilePath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlFilePath}'.");
                return;
            }

            // Load the HTML content into a new workbook instance
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // OPTIONAL: List all imported hyperlinks using the worksheet's Hyperlinks collection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Hyperlink hyperlink in sheet.Hyperlinks)
                {
                    // Get the cell address of the hyperlink start position
                    CellArea area = hyperlink.Area;
                    string cellName = sheet.Cells[area.StartRow, area.StartColumn].Name;
                    string address = hyperlink.Address ?? string.Empty;
                    Console.WriteLine($"Hyperlink found at {sheet.Name}!{cellName}: {address}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an XLSX file
            workbook.Save(excelFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"HTML has been converted to Excel and saved as '{excelFilePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
