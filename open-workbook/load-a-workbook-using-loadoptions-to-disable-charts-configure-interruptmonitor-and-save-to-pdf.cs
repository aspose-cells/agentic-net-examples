// Title: Load an XLSX workbook with LoadOptions, assign an InterruptMonitor, and save as PDF using Aspose.Cells for .NET
// Common Searches: how to load an Excel file with LoadOptions in Aspose.Cells C# | assign an InterruptMonitor to a Workbook before exporting to PDF with Aspose.Cells | export an XLSX workbook to PDF using Aspose.Cells while monitoring interruptions | Aspose.Cells C# example for loading workbook and saving as PDF | check file existence before loading workbook with Aspose.Cells
// Tags: load workbook using LoadOptions Aspose.Cells | configure InterruptMonitor for Aspose.Cells workbook | save workbook as PDF Aspose.Cells | Aspose.Cells .NET Excel to PDF conversion | exception handling in Aspose.Cells file operations

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input XLSX file, loads it into a Workbook with LoadOptions, attaches an InterruptMonitor to control execution, and then saves the workbook as a PDF, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with default options (LoadDataOnly not available in this version)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Configure an interrupt monitor (no interruption)
            InterruptMonitor monitor = new InterruptMonitor();
            workbook.InterruptMonitor = monitor;

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
