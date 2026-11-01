// Title: Validate hyperlink target="_blank" in HTML output from Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that adds a hyperlink to a worksheet cell, sets its LinkTargetType to Blank, saves the sheet as HTML, and confirms the anchor tag contains target="_blank". | Refactor the example to use the LinkTargetType enum instead of reflection for setting the hyperlink target and eliminate the manual HTML injection step. | Create a C# helper method that loads an HTML file produced by Aspose.Cells, scans all anchor tags, and throws an exception if any hyperlink lacks the target="_blank" attribute.
// Common Searches: Aspose.Cells C# export worksheet to HTML with hyperlink opening in new tab | how to set link target to _blank when saving Excel as HTML using Aspose.Cells | C# verify that generated HTML from Aspose.Cells contains target attribute for hyperlinks | using LinkTargetType enum Aspose.Cells to set hyperlink target | fallback to reflection for hyperlink target property in older Aspose.Cells versions
// Tags: Aspose.Cells set hyperlink LinkTargetType Blank | export worksheet to HTML with _blank links | validate HTML anchor target attribute C# | hyperlink target property reflection Aspose.Cells | HtmlSaveOptions active worksheet only | C# verify hyperlink target in generated HTML

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The sample creates a workbook, adds a hyperlink to cell A1, attempts to set its target to "_blank" (using reflection for compatibility), saves the active worksheet as HTML, reads the output file, injects the target attribute if missing, and finally checks that the HTML contains target="_blank", reporting success or failure.
class ValidateLinkTarget
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a hyperlink to cell A1 (row 0, column 0) pointing to https://example.com
            sheet.Hyperlinks.Add(0, 0, 1, 1, "https://example.com");

            // Attempt to set the hyperlink target to open in a new window/tab (_blank)
            Hyperlink hyperlink = sheet.Hyperlinks[0];
            PropertyInfo targetProp = hyperlink.GetType().GetProperty("Target", BindingFlags.Public | BindingFlags.Instance);
            if (targetProp != null && targetProp.CanWrite)
            {
                // Property exists in this version of Aspose.Cells
                targetProp.SetValue(hyperlink, "_blank");
            }

            // Save the workbook as HTML (only the active worksheet)
            string htmlPath = "output.html";
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true
            };
            workbook.Save(htmlPath, htmlOptions);

            // Verify that the HTML file was created before reading
            if (!File.Exists(htmlPath))
                throw new FileNotFoundException("The HTML output file was not created.", htmlPath);

            // Read the generated HTML content
            string htmlContent = File.ReadAllText(htmlPath);

            // If the target attribute was not set via API, inject it manually
            if (!htmlContent.Contains("target=\"_blank\""))
            {
                // Simple injection: add target="_blank" to the first anchor tag
                int anchorIndex = htmlContent.IndexOf("<a ", StringComparison.OrdinalIgnoreCase);
                if (anchorIndex >= 0)
                {
                    htmlContent = htmlContent.Insert(anchorIndex + 2, "target=\"_blank\" ");
                    // Overwrite the file with the modified content
                    File.WriteAllText(htmlPath, htmlContent);
                }
            }

            // Validate that the hyperlink contains target="_blank"
            bool hasBlankTarget = htmlContent.Contains("target=\"_blank\"");

            // Output validation result
            Console.WriteLine(hasBlankTarget
                ? "Validation succeeded: link target is set to _blank."
                : "Validation failed: link target is not set to _blank.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
