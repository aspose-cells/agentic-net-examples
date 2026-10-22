// Title: Export an Excel worksheet to HTML with visible gridlines using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, populates cells, and saves it as HTML with gridlines enabled via Aspose.Cells HtmlSaveOptions. | Update existing Aspose.Cells C# code to turn on ExportGridLines so the generated HTML file shows cell borders.
// Common Searches: Aspose.Cells C# export to HTML with gridlines visible | HtmlSaveOptions ExportGridLines property example .NET | Why are gridlines missing in HTML output from Aspose.Cells | Enable cell borders when saving workbook as HTML using Aspose.Cells | C# code to save Excel as HTML with visible gridlines
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# export workbook to HTML with gridlines | Aspose.Cells HTML gridlines visibility | Enable cell borders in Aspose.Cells HTML output | Export Excel to HTML using Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a new Workbook, fills cells A1‑B4 with sample data, configures HtmlSaveOptions with ExportGridLines = true, and saves the workbook as an HTML file where the gridlines are rendered.
class Program
{
    static void Main()
    {
        // Create a new workbook
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
        sheet.Cells["A4"].PutValue("Cherry");
        sheet.Cells["B4"].PutValue(2.50);

        // Set gridlines to be visible in the HTML output
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportGridLines = true   // Enable gridlines
        };

        // Save the workbook as an HTML file with gridlines
        workbook.Save("ExportedWithGridLines.html", saveOptions);
    }
}
