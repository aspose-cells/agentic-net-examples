// Title: Export an Excel workbook to HTML in C# with Aspose.Cells while disabling tooltip text for faster rendering
// AI Prompts: Write C# code that saves a Workbook as an HTML file using Aspose.Cells and sets HtmlSaveOptions.AddTooltipText to false. | Show how to configure Aspise.Cells HtmlSaveOptions to turn off tooltip generation and improve HTML export speed. | Create a minimal C# example that populates a worksheet and exports it to HTML without cell tooltips.
// Common Searches: Aspose.Cells C# HtmlSaveOptions AddTooltipText false performance | how to disable cell tooltips when exporting Excel to HTML with Aspose.Cells | speed up HTML export of large workbook using Aspose.Cells | C# export workbook to HTML without tooltip text Aspose | HtmlSaveOptions settings to improve rendering speed Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions tooltip off | C# export workbook to HTML fast | HTML export performance Aspose.Cells | Excel to HTML no tooltips | Aspose.Cells rendering speed optimization

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The program creates a workbook, fills it with sample data, configures HtmlSaveOptions with AddTooltipText set to false to eliminate tooltip generation, and saves the result as an HTML file, resulting in noticeably faster rendering.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (you can also load an existing one using new Workbook("input.xlsx"))
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Disable tooltip text generation to improve rendering speed
                AddTooltipText = false
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save("ExportedReport.html", htmlOptions);
        }
    }
}
