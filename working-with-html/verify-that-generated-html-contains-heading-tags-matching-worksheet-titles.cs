// Title: Check that each worksheet name appears inside an <h1‑h6> tag after exporting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: C# how to confirm Excel sheet names are included in heading tags after Aspose.Cells HTML export | using Aspose.Cells to export workbook to single HTML and verify worksheet titles with regex | regex pattern to find worksheet name inside h1‑h6 tags in Aspose.Cells generated HTML | Aspose.Cells .NET export all worksheets to HTML and validate heading elements | automate verification of worksheet titles in HTML output using Aspose.Cells C#
// Tags: verify worksheet titles in HTML using Aspose.Cells | Aspose.Cells HTML export heading validation | regex worksheet name detection in exported HTML | export workbook to single HTML with Base64 images Aspose.Cells | C# check Excel sheet headings after HTML conversion

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, saves it to a single HTML string with HtmlSaveOptions (Base64‑encoded images and all worksheets), then iterates each worksheet and uses a case‑insensitive regex to locate the sheet name inside any <h1>‑<h6> tag. It records pass/fail results for each sheet and prints a verification summary to the console.
class HtmlWorksheetHeadingVerifier
{
    static void Main()
    {
        // Path to the source Excel file
        string excelPath = "input.xlsx";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(excelPath);

        // Export the workbook to HTML and keep the result in a string
        string htmlContent;
        using (MemoryStream htmlStream = new MemoryStream())
        {
            // SaveOptions for HTML export (lifecycle rule: save)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Export all worksheets, not only the active one
                ExportActiveWorksheetOnly = false,
                // Generate a single HTML file (no separate images folder)
                ExportImagesAsBase64 = true
            };

            workbook.Save(htmlStream, saveOptions);
            htmlStream.Position = 0;
            using (StreamReader reader = new StreamReader(htmlStream))
            {
                htmlContent = reader.ReadToEnd();
            }
        }

        // Prepare a list to collect verification results
        List<string> verificationResults = new List<string>();

        // Iterate through each worksheet and verify its title appears in a heading tag
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            string sheetName = sheet.Name;

            // Regex to find any heading tag (h1-h6) that contains the sheet name (case‑insensitive)
            string pattern = $@"<h[1-6][^>]*>\s*{Regex.Escape(sheetName)}\s*</h[1-6]>";
            bool found = Regex.IsMatch(htmlContent, pattern, RegexOptions.IgnoreCase);

            verificationResults.Add(found
                ? $"PASS: Worksheet \"{sheetName}\" found in a heading tag."
                : $"FAIL: Worksheet \"{sheetName}\" NOT found in any heading tag.");
        }

        // Output the verification summary
        Console.WriteLine("HTML Worksheet Heading Verification Results:");
        foreach (string result in verificationResults)
        {
            Console.WriteLine(result);
        }
    }
}
