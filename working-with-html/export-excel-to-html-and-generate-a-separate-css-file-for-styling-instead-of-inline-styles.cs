// Title: Export Excel to HTML with an external CSS file using Aspose.Cells for .NET (C#)
// Description: Shows how to load an XLSX workbook with Aspose.Cells, enable HtmlSaveOptions.ExportWorksheetCSSSeparately, and save the workbook as HTML so that all formatting is written to a separate CSS file in the same output folder.
// Keywords: Aspose.Cells C# HTML export | ExportWorksheetCSSSeparately | Excel to HTML external CSS | Aspose.Cells save workbook as HTML | separate stylesheet Aspose.Cells | C# generate CSS file from Excel
// Common Searches: Aspose.Cells export Excel to HTML separate CSS | C# HtmlSaveOptions ExportWorksheetCSSSeparately example | How to generate external stylesheet when saving workbook as HTML | Save workbook as HTML with CSS file Aspose.Cells .NET
// Developer Intent: Create an HTML version of an Excel workbook and output all styling to a distinct CSS file instead of inline styles.
// Use Cases: Publish financial or analytical reports on a website while keeping the style definitions in a reusable stylesheet. | Integrate Excel‑derived tables into a web application where CSS must be managed centrally for consistent theming. | Automate batch conversion of multiple workbooks to HTML with shared external CSS to simplify maintenance and reduce page size.
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as HTML with the CSS exported to a separate file. | Explain the purpose of the ExportWorksheetCSSSeparately property and how to reference the generated CSS file in the resulting HTML page. | Show how to customize the filename and location of the external CSS file when exporting Excel to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an XLSX workbook with Aspose.Cells, enable HtmlSaveOptions.ExportWorksheetCSSSeparately, and save the workbook as HTML so that all formatting is written to a separate CSS file in the same output folder.
class ExportExcelToHtmlWithSeparateCss
{
    static void Main()
    {
        // Load an existing workbook (provide the correct path to your Excel file)
        string inputPath = "sample.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options to generate a separate CSS file
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Ensure the output directory exists
        string outputDir = "HtmlOutput";
        Directory.CreateDirectory(outputDir);

        // Save the workbook as HTML; the CSS will be written to a separate file in the same folder
        string htmlPath = Path.Combine(outputDir, "sample.html");
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine("HTML file saved to: " + htmlPath);
        Console.WriteLine("Separate CSS file generated alongside the HTML.");
    }
}
