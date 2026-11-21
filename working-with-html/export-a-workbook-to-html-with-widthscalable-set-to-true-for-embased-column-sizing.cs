// Title: Export an Excel workbook to HTML with column widths expressed in em units using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML with WidthScalable true | how to generate HTML from .xlsx with scalable column widths in Aspose.Cells | set column width units to em when saving Excel as HTML using Aspose.Cells .NET | HtmlSaveOptions WidthScalable example for C# Aspose.Cells | convert Excel file to HTML preserving column width scaling in .NET
// Tags: HtmlSaveOptions WidthScalable setting | C# Excel to HTML conversion Aspose.Cells | em based column width scaling | Aspose.Cells HTML export options | save workbook as HTML with scalable width

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, configures HtmlSaveOptions with WidthScalable set to true so column widths are output in em units, and saves the workbook as an HTML file, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable scalable width so column widths are expressed in em units
                WidthScalable = true
                // Note: ColumnWidthType property may not be available in older versions;
                // if needed, it can be set here when supported.
            };

            // Export the workbook to HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
