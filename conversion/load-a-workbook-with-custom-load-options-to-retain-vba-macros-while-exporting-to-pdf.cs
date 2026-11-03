// Title: Load an .xlsm workbook with VBA macros and save it as PDF using Aspose.Cells for .NET
// Common Searches: how to open an .xlsm file with Aspose.Cells in C# and convert it to PDF | C# Aspose.Cells preserve VBA macros when converting Excel to PDF | load macro-enabled workbook using Aspose.Cells and export to PDF without losing macros | Aspose.Cells .NET convert macro-enabled Excel workbook to PDF example
// Tags: Aspose.Cells .xlsm loading | Aspose.Cells PDF export | C# preserve VBA macros during conversion | Aspose.Cells custom load options | macro-enabled Excel PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an input.xlsm file, loads it into an Aspose.Cells Workbook (keeping VBA macros in memory), and then saves the workbook as a PDF file using SaveFormat.Pdf, with error handling for any runtime issues.
class Program
{
    static void Main()
    {
        // Path to the source workbook that contains VBA macros
        string sourcePath = "input.xlsm";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: The file \"{sourcePath}\" was not found.");
            return;
        }

        // Path for the output PDF
        string pdfPath = "output.pdf";

        try
        {
            // Load the workbook (Aspose.Cells automatically detects the format)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as PDF; macros are retained in the workbook object but not in the PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at \"{pdfPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
