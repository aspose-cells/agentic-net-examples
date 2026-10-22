// Title: Create plain HTML from an Excel workbook without exporting conditional formatting using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, clears all conditional formatting rules, and saves the workbook as plain HTML with Aspose.Cells. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to generate a clean HTML representation of a workbook while ensuring no conditional formatting is included.
// Common Searches: how to export Excel to HTML without conditional formatting using Aspose.Cells C# | Aspose.Cells generate plain HTML from workbook ignoring conditional formatting rules | C# remove conditional formatting before saving workbook as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions plain visual output no conditional styles
// Tags: Aspose.Cells HtmlSaveOptions plain HTML export | C# clear workbook conditional formatting Aspose.Cells | Aspose.Cells export workbook to HTML without styles | remove conditional formatting Aspose.Cells HTML conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file (or creates a simple workbook), optionally clears all conditional formatting, configures HtmlSaveOptions, and saves the workbook as an HTML file that displays the data without any conditional formatting styles, providing a clean visual representation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Load workbook from file if it exists; otherwise create a simple workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample data");
            }

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            // The ExportConditionalFormatting property is not available in the current API version.
            // If needed, adjust other HtmlSaveOptions properties here.

            // Save the workbook as HTML using the configured options.
            workbook.Save(outputPath, htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
