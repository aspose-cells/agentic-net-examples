// Title: Convert an HTML file to an Excel workbook with clickable hyperlinks using Aspose.Cells for .NET
// AI Prompts: Load an HTML document into an Aspose.Cells Workbook while preserving <a> tag hyperlinks. | Iterate through each worksheet's Hyperlink collection to log the address of every link before saving. | Save the workbook as an XLSX file ensuring that all hyperlinks remain active in Excel. | Create the output folder if it does not exist and handle missing input file errors gracefully.
// Common Searches: Aspose.Cells C# convert HTML to XLSX with hyperlinks intact | how to keep anchor tags clickable when exporting HTML to Excel using Aspose | load HTML into workbook preserving hyperlinks Aspose.Cells .NET | C# example for converting HTML file to Excel while retaining links | Aspose.Cells hyperlink extraction from HTML before saving to XLSX
// Tags: HTML to XLSX conversion preserving hyperlinks Aspose.Cells | load HTML workbook with active links C# | enumerate worksheet hyperlinks Aspose.Cells | save workbook as XLSX clickable links | load options HTML format Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates loading an HTML file into an Aspose.Cells Workbook using LoadOptions for HTML, iterating the Hyperlinks collection to verify addresses, and saving the workbook as an XLSX file so that all original <a> tags become clickable links in Excel.
class Program
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Path where the resulting Excel file will be saved
        string excelFilePath = "output.xlsx";

        // Verify that the input HTML file exists
        if (!File.Exists(htmlFilePath))
        {
            Console.WriteLine($"Input file not found: {htmlFilePath}");
            return;
        }

        try
        {
            // Load the HTML file into a Workbook while preserving hyperlinks
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFilePath, loadOptions);

            // Iterate through hyperlinks to confirm they are loaded correctly
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Hyperlink link in sheet.Hyperlinks)
                {
                    // Display the hyperlink address; cell position properties are not required for core functionality
                    Console.WriteLine($"Hyperlink found: Address = {link.Address}");
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(excelFilePath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an XLSX file; hyperlinks remain clickable in the workbook
            workbook.Save(excelFilePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {excelFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
