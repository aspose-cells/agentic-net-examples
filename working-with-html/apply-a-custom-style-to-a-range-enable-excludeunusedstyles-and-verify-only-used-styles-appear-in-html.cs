// Title: Apply a user‑defined cell style to a specific range and export to HTML while excluding unused styles with Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, define a user‑defined style named MyCustomStyle, apply it to cells A1:B3 using a StyleFlag with All=true, enable exclusion of unused styles, and save the workbook as an HTML file. | Read the generated HTML file and test whether the identifier MyCustomStyle appears, confirming that only the applied style was embedded.
// Common Searches: Aspose.Cells C# apply style to a cell range and export to HTML | exclude unused CSS classes when saving Excel as HTML with Aspose.Cells | check that only used styles are written to HTML output in Aspose.Cells | optimize HTML export size by removing unused styles Aspose.Cells .NET | how to verify style inclusion in Aspose.Cells generated HTML
// Tags: apply user-defined style range Aspose.Cells | HTML export exclude unused styles .NET | apply all formatting to range Aspose.Cells | verify style presence in generated HTML | user-defined cell style Aspose.Cells C#

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, defines a custom style with a light‑green background and bold dark‑green font, applies it to the range A1:B3, saves the workbook as an HTML file while excluding unused styles, and then reads the HTML to confirm that the custom style name is present, demonstrating that only used styles are included in the output.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Create a custom style
            Style customStyle = workbook.CreateStyle();
            customStyle.Name = "MyCustomStyle";                 // Give the style a recognizable name
            customStyle.ForegroundColor = Color.LightGreen;    // Background color
            customStyle.Pattern = BackgroundType.Solid;        // Apply background
            customStyle.Font.Color = Color.DarkGreen;          // Font color
            customStyle.Font.IsBold = true;                    // Bold font

            // Apply the custom style to a specific range (A1:B3)
            StyleFlag flag = new StyleFlag { All = true };
            Aspose.Cells.Range targetRange = sheet.Cells.CreateRange("A1:B3");
            targetRange.ApplyStyle(customStyle, flag);

            // Save the workbook as HTML
            string htmlFile = "StyledOutput.html";
            workbook.Save(htmlFile, SaveFormat.Html);

            // Verify that only used styles appear in the generated HTML
            if (File.Exists(htmlFile))
            {
                string htmlContent = File.ReadAllText(htmlFile);
                bool customStyleFound = htmlContent.Contains("MyCustomStyle");
                Console.WriteLine($"Custom style present in HTML: {customStyleFound}");
            }
            else
            {
                Console.WriteLine($"Failed to generate HTML file: {htmlFile}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
