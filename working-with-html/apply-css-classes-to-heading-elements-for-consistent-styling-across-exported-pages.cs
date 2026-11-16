// Title: Apply a named CSS class to heading cells and export the worksheet as HTML using Aspose.Cells for .NET
// Common Searches: how to generate a CSS class from a style name when saving Excel to HTML with Aspose.Cells C# | Aspose.Cells export worksheet to HTML with custom heading style applied to cells | C# code to assign a CSS class to specific cells for HTML output using Aspose.Cells | save Excel workbook as HTML with embedded CSS for headings using Aspose.Cells .NET | set style name to create CSS class for heading rows in Aspose.Cells HTML export
// Tags: named CSS class generation Aspose.Cells | heading cell style application C# | Aspose.Cells HTML save options embedded CSS | style name to CSS class mapping Aspose.Cells | export workbook with custom heading formatting HTML

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, defines a bold blue 14‑point style named "headingStyle", applies it to cells A1 and A2, and saves the active worksheet as an HTML file using HtmlSaveOptions, where the style name is emitted as a CSS class for consistent heading styling.
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

            // Add heading texts to cells
            sheet.Cells["A1"].PutValue("Sales Report");
            sheet.Cells["A2"].PutValue("Quarterly Summary");

            // Create a style that will be exported as a CSS class (if supported)
            Style headingStyle = workbook.CreateStyle();
            headingStyle.Font.IsBold = true;
            headingStyle.Font.Size = 14;
            headingStyle.Font.Color = Color.Blue;
            // Assign a name – this becomes the CSS class name in the generated CSS (when supported)
            headingStyle.Name = "headingStyle";

            // Apply the style to the heading cells
            sheet.Cells["A1"].SetStyle(headingStyle);
            sheet.Cells["A2"].SetStyle(headingStyle);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true
                // Note: In this version of Aspose.Cells, external CSS generation properties
                // (CssStyleSheetType, CssFileName) are not available. The HTML will contain
                // embedded CSS, which still preserves the defined style.
            };

            // Save the workbook as HTML
            workbook.Save("Report.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
