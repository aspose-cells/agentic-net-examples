// Title: Convert HTML with Hyperlinks to Clickable Excel (XLSX) using Aspose.Cells for .NET
// Description: Load a local HTML file that contains anchor tags into an Aspose.Cells Workbook, automatically preserve each hyperlink, and save the workbook as an XLSX file where the links remain active. Includes optional code to enumerate and log imported hyperlinks and basic error handling for missing files.
// Keywords: Aspose.Cells HTML to Excel | preserve hyperlinks C# | convert HTML to XLSX .NET | load HTML workbook Aspose | hyperlink collection Aspose.Cells | clickable links Excel export
// Common Searches: Aspose.Cells keep hyperlinks when converting HTML to Excel | C# convert HTML file to XLSX with active links | load HTML into Workbook and export clickable hyperlinks | iterate over imported hyperlinks Aspose.Cells | error handling missing HTML file Aspose.Cells
// Developer Intent: Import an HTML document with anchor tags into a Workbook and export it as an Excel file that retains functional hyperlinks.
// Use Cases: Validate the existence of the source HTML file before conversion to prevent runtime errors. | Load the HTML file into a Workbook; Aspose.Cells automatically creates Hyperlink objects for each <a> tag. | Save the Workbook as XLSX so the hyperlinks are clickable in Excel. | Optionally iterate through Worksheet.Hyperlinks to log, audit, or modify link addresses before saving.
// AI Prompts: Write C# code that uses Aspose.Cells to convert an HTML file with embedded <a> tags into an XLSX workbook, ensuring all links stay clickable. | Explain how to access the Hyperlink collection after loading HTML, and show how to log or update hyperlink URLs. | Provide best‑practice error handling for missing HTML input and verification that hyperlinks were imported correctly.

using System;
using System.IO;
using Aspose.Cells;

// Load a local HTML file that contains anchor tags into an Aspose.Cells Workbook, automatically preserve each hyperlink, and save the workbook as an XLSX file where the links remain active. Includes optional code to enumerate and log imported hyperlinks and basic error handling for missing files.
class HtmlToExcel
{
    static void Main()
    {
        // Input HTML file containing hyperlinks
        string htmlPath = "sample.html";

        // Output Excel file where hyperlinks will be clickable
        string excelPath = "output.xlsx";

        // Verify that the input HTML file exists to avoid FileNotFoundException
        if (!File.Exists(htmlPath))
        {
            Console.WriteLine($"Error: The HTML file \"{htmlPath}\" was not found.");
            return;
        }

        try
        {
            // Load the HTML file into a workbook (hyperlinks are preserved automatically)
            Workbook workbook = new Workbook(htmlPath);

            // Optional: iterate through hyperlinks to verify they were imported
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Hyperlink link in sheet.Hyperlinks)
                {
                    // Display hyperlink address (row/column properties may not be available in all versions)
                    Console.WriteLine($"Hyperlink -> {link.Address}");
                }
            }

            // Save the workbook as an Excel file (XLSX)
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine("HTML successfully converted to Excel with clickable hyperlinks.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
