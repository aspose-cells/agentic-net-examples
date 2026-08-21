// Title: Convert Excel to HTML with Aspose.Cells (C#) – include hidden worksheets
// Description: Loads an .xlsx file into an Aspose.Cells Workbook, creates HtmlSaveOptions with default settings, sets ExportHiddenWorksheet to true, and saves the workbook as an HTML file. All worksheets, including hidden ones, are rendered in the output.
// Keywords: Aspose.Cells | C# Excel to HTML | HtmlSaveOptions | ExportHiddenWorksheet | convert hidden worksheets to HTML | default HTML conversion | Workbook.Save HTML | Aspose.Cells .NET
// Common Searches: Aspose.Cells export hidden worksheets to HTML C# | Convert Excel workbook to HTML with default options | HtmlSaveOptions ExportHiddenWorksheet example | C# save .xlsx as .html using Aspose.Cells | How to include hidden sheets in HTML export Aspose
// Developer Intent: The developer needs to transform an Excel workbook into an HTML document while ensuring that any hidden worksheets are also rendered.
// Use Cases: Publish a complete Excel report on a website, showing hidden analysis sheets in HTML. | Create HTML snapshots of workbooks for email or documentation, preserving all worksheet visibility. | Automate batch conversion of multiple .xlsx files to .html for an internal knowledge base, using default settings and exporting hidden sheets.
// AI Prompts: Generate C# code with Aspose.Cells to convert an Excel file to HTML and include hidden worksheets. | Explain the default behavior of HtmlSaveOptions.ExportHiddenWorksheet and when to set it explicitly. | Provide a script that batch processes a folder of Excel files, converting each to HTML while exporting hidden worksheets.

using System;
using Aspose.Cells;

// Loads an .xlsx file into an Aspose.Cells Workbook, creates HtmlSaveOptions with default settings, sets ExportHiddenWorksheet to true, and saves the workbook as an HTML file. All worksheets, including hidden ones, are rendered in the output.
class Program
{
    static void Main()
    {
        // Source Excel file path
        string sourcePath = "input.xlsx";

        // Destination HTML file path
        string outputPath = "output.html";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options with default settings
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Explicitly export hidden worksheets (default is true, set for clarity)
        saveOptions.ExportHiddenWorksheet = true;

        // Save the workbook as an HTML file using the specified options
        workbook.Save(outputPath, saveOptions);
    }
}
