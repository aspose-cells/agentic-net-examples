// Title: Set hyperlink target to _blank and verify in HTML export using Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, add a hyperlink to cell A1 with display text, set its Target property to "_blank" via reflection, save the workbook as HTML, and output the generated file path. | Insert a hyperlink that opens in a new browser tab using Aspose.Cells, export the worksheet to HTML, then read the HTML file to confirm the anchor includes target="_blank". | Write C# code that adds a hyperlink, assigns a new‑window target, saves the workbook as HTML, and programmatically asserts that the exported HTML contains the _blank target attribute.
// Common Searches: Aspose.Cells .NET how to make hyperlink open in new tab when saving as HTML | set hyperlink target attribute _blank using Aspose.Cells workbook | verify that exported HTML from Aspose.Cells contains target="_blank" | use reflection to set hyperlink Target property in Aspose.Cells C# | Aspose.Cells hyperlink export HTML new window link
// Tags: aspocells hyperlink target blank | aspocells export html with link target | c# reflection set hyperlink target aspocells | verify html anchor target aspocells | hyperlink new tab export aspocells

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The example creates a workbook, adds a hyperlink to cell A1, uses reflection to set its Target property to "_blank", saves the workbook as HTML, and then reads the generated file to confirm that the hyperlink markup includes target="_blank".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define hyperlink parameters
            int row = 0;               // zero‑based index for row 1 (A)
            int column = 0;            // zero‑based index for column A
            string url = "https://example.com";
            string displayText = "Example Link";

            // Add a hyperlink to a single cell (row, column, 1 row, 1 column)
            int hyperlinkIndex = sheet.Hyperlinks.Add(row, column, 1, 1, url);
            Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
            hyperlink.TextToDisplay = displayText;

            // Set target to open in a new window/tab if the property exists (via reflection)
            PropertyInfo targetProp = hyperlink.GetType().GetProperty("Target");
            if (targetProp != null && targetProp.CanWrite)
            {
                targetProp.SetValue(hyperlink, "_blank", null);
            }

            // Prepare output path and ensure directory exists
            string htmlPath = "output.html";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(htmlPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML
            workbook.Save(htmlPath, SaveFormat.Html);

            // Verify that the HTML file was created and contains the expected target attribute
            if (File.Exists(htmlPath))
            {
                string htmlContent = File.ReadAllText(htmlPath);
                bool containsBlankTarget = htmlContent.Contains("target=\"_blank\"");
                Console.WriteLine(containsBlankTarget
                    ? "Validation passed: link target is \"_blank\"."
                    : "Validation failed: link target \"_blank\" not found.");
            }
            else
            {
                Console.WriteLine($"Error: The file '{htmlPath}' was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
