// Title: Generate HTML for only the active worksheet using Aspose.Cells HtmlSaveOptions in C#
// AI Prompts: Write C# code that loads an Excel workbook, activates a chosen worksheet, sets HtmlSaveOptions.ExportActiveWorksheetOnly to true, and saves the result as an HTML file. | Adapt existing Aspose.Cells code so that only the currently active sheet is exported to HTML by configuring HtmlSaveOptions appropriately.
// Common Searches: Aspose.Cells C# export only the active worksheet to HTML file | How to use HtmlSaveOptions.ExportActiveWorksheetOnly with Aspose.Cells .NET | Save a single sheet from an Excel workbook as HTML using Aspose.Cells | C# example converting the active worksheet to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly C# | active worksheet html export Aspose.Cells | single sheet to html conversion Aspose.Cells | csharp excel to html single worksheet

using Aspose.Cells;

// The sample loads an Excel workbook, enables HtmlSaveOptions.ExportActiveWorksheetOnly, and saves the workbook as an HTML file that contains only the active worksheet.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure the desired worksheet is the active one (optional)
        // workbook.Worksheets[0].Activate();

        // Configure HTML save options to export only the active worksheet
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportActiveWorksheetOnly = true;

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
