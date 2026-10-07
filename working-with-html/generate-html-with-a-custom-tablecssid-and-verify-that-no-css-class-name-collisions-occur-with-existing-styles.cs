// Title: Export an Aspose.Cells workbook to HTML with a unique TableCssId and detect CSS class collisions in C#
// AI Prompts: Create C# code that saves a Workbook as HTML using HtmlSaveOptions, assigns a GUID‑based TableCssId, and checks the generated HTML for any CSS class with the same name. | Modify the export to prepend a custom prefix to the TableCssId and log a warning if that prefixed ID already exists among CSS class selectors. | Implement a regex that extracts all CSS class selectors from the HTML stream and validates that the TableCssId is not present.
// Common Searches: Aspose.Cells C# how to set a custom TableCssId when saving to HTML | detecting CSS id and class name collisions in Aspose.Cells HTML output | generate unique table identifier for Excel to HTML conversion using Aspose.Cells | C# regex to list CSS class selectors from Aspose.Cells generated HTML | prevent TableCssId conflict with existing CSS classes in exported HTML
// Tags: Aspose.Cells HTML export custom TableCssId | C# detect CSS id class collision Aspose.Cells | unique table identifier GUID C# | regex extract CSS class selectors Aspose HTML | memory stream HTML save options Aspose.Cells

using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Aspose.Cells;

// The program creates a workbook, adds sample data, generates a GUID‑based TableCssId, saves the workbook as HTML with HtmlSaveOptions, extracts CSS class names via a regular expression, and verifies that the TableCssId does not collide with any existing CSS class in the output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header1");
            sheet.Cells["B1"].PutValue("Header2");
            sheet.Cells["A2"].PutValue("Data1");
            sheet.Cells["B2"].PutValue("Data2");

            // Generate a unique TableCssId to avoid collisions
            string tableCssId = "tbl_" + Guid.NewGuid().ToString("N");

            // Configure HTML save options with the custom TableCssId
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // export only the active sheet
                TableCssId = tableCssId            // set custom table id
            };

            // Save the workbook to a memory stream as HTML
            using (MemoryStream htmlStream = new MemoryStream())
            {
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0;

                string htmlContent;
                using (StreamReader reader = new StreamReader(htmlStream))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Extract all CSS class names defined in the generated HTML
                HashSet<string> cssClassNames = new HashSet<string>();
                Regex classRegex = new Regex(@"\.([A-Za-z0-9_-]+)\s*\{", RegexOptions.Compiled);
                foreach (Match match in classRegex.Matches(htmlContent))
                {
                    cssClassNames.Add(match.Groups[1].Value);
                }

                // Verify that the TableCssId does not collide with any CSS class name
                if (cssClassNames.Contains(tableCssId))
                {
                    Console.WriteLine("Collision detected: TableCssId matches an existing CSS class name.");
                }
                else
                {
                    Console.WriteLine($"No collision detected. TableCssId = \"{tableCssId}\"");
                }

                // Optionally write the HTML to a file for inspection
                File.WriteAllText("output.html", htmlContent);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
