// Title: Create a Single‑Worksheet MHT File from Excel with Aspose.Cells for .NET (C#)
// Description: Opens an existing Excel workbook, selects the first worksheet as active, configures HtmlSaveOptions for MHTML output with ExportActiveWorksheetOnly enabled, and saves the result as a single‑worksheet MHT file.
// Keywords: Aspose.Cells | C# | .NET | MHT | MHTML | HtmlSaveOptions | ExportActiveWorksheetOnly | single worksheet export | Excel to MHT conversion | save workbook as MHT
// Common Searches: Aspose.Cells save active sheet as MHT | C# export single worksheet to MHTML | How to generate MHT from Excel using Aspose | HtmlSaveOptions ExportActiveWorksheetOnly example | Convert Excel workbook to MHT file .NET
// Developer Intent: Generate an MHT document that contains only the active worksheet of an existing Excel file.
// Use Cases: Attach a specific worksheet snapshot to an email in a web‑friendly format. | Publish a dashboard sheet on an intranet portal without exposing the whole workbook. | Archive a single worksheet as a self‑contained HTML file for documentation.
// AI Prompts: Write C# code that saves the third worksheet of an Excel file as a separate MHT file using Aspose.Cells. | Show how to embed images and apply a custom CSS stylesheet when exporting a worksheet to MHTML with Aspose.Cells. | Provide a loop that iterates through selected worksheets and creates individual MHT files for each.

using System;
using Aspose.Cells;

// Opens an existing Excel workbook, selects the first worksheet as active, configures HtmlSaveOptions for MHTML output with ExportActiveWorksheetOnly enabled, and saves the result as a single‑worksheet MHT file.
class Program
{
    static void Main()
    {
        // Path to the source workbook (replace with your actual file)
        string sourcePath = "input.xlsx";

        // Open the existing workbook
        Workbook workbook = new Workbook(sourcePath);

        // Set the worksheet you want to export as the active sheet
        workbook.Worksheets.ActiveSheetIndex = 0; // first worksheet

        // Configure save options for MHTML and export only the active worksheet
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHtml);
        saveOptions.ExportActiveWorksheetOnly = true; // single worksheet

        // Save the workbook as a single‑worksheet MHT file
        string outputPath = "single_sheet.mht";
        workbook.Save(outputPath, saveOptions);
    }
}
