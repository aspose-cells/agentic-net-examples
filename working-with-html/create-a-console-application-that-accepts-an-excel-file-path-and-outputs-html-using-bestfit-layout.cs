// Title: C# Console App – Convert Excel to Best‑Fit HTML with Aspose.Cells
// Description: A .NET console program that validates a file path, loads an Excel workbook with Aspose.Cells, applies HtmlSaveOptions (LayoutMode = Normal, ExportDataOptions = All) and saves the sheet as a best‑fit HTML file.
// Keywords: Aspose.Cells | C# convert Excel to HTML | HtmlSaveOptions LayoutMode Normal | best fit HTML layout | ExportDataOptions.All | .NET console Excel to HTML | Excel workbook to web page
// Common Searches: Aspose.Cells C# example for Excel to HTML conversion | How to use HtmlSaveOptions LayoutMode Normal in .NET | Console application that exports Excel as HTML | Best‑fit HTML layout for Excel files using Aspose | Save Excel workbook to HTML with all data in C#
// Developer Intent: Create a command‑line tool that reads an .xlsx file and outputs a best‑fit HTML representation using Aspose.Cells.
// Use Cases: Publish financial or sales reports as static HTML pages without Office installed. | Batch‑process a folder of Excel files for intranet documentation. | Provide a lightweight preview utility for end‑users to view spreadsheets in a browser.
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, sets HtmlSaveOptions.LayoutMode to Normal, enables ExportDataOptions.All, and saves the result as HTML. | Explain step‑by‑step how to build a .NET console app that accepts an Excel file path argument and creates a best‑fit HTML file. | Suggest modifications to handle multiple input files and custom output directories in the console converter.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A .NET console program that validates a file path, loads an Excel workbook with Aspose.Cells, applies HtmlSaveOptions (LayoutMode = Normal, ExportDataOptions = All) and saves the sheet as a best‑fit HTML file.
class Program
{
    static void Main(string[] args)
    {
        // Verify arguments
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: <exe> <excelFilePath> [outputHtmlPath]");
            return;
        }

        string excelPath = args[0];
        if (!System.IO.File.Exists(excelPath))
        {
            Console.WriteLine($"File not found: {excelPath}");
            return;
        }

        // Determine output HTML path
        string htmlPath = args.Length > 1
            ? args[1]
            : System.IO.Path.ChangeExtension(excelPath, ".html");

        try
        {
            // Load the workbook from the provided Excel file
            Workbook workbook = new Workbook(excelPath);

            // Create HTML save options (constructor rule)
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set layout mode to Normal for best‑fit rendering (property rule)
            saveOptions.LayoutMode = HtmlLayoutMode.Normal;

            // Export all data (property rule)
            saveOptions.ExportDataOptions = HtmlExportDataOptions.All;

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {htmlPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
