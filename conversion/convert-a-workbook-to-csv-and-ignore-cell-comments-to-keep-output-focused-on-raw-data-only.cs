// Title: Convert an Excel workbook to CSV while excluding cell comments using Aspose.Cells for .NET
// Common Searches: how to export .xlsx to .csv without comments using Aspose.Cells C# | Aspose.Cells C# save workbook as CSV ignore cell comments | C# TxtSaveOptions ExportCellComments false example | convert Excel file to CSV with Aspose.Cells while dropping comments | Aspose.Cells .NET CSV export excluding comments from cells | save workbook as CSV in C# using Aspose.Cells TxtSaveOptions
// Tags: Aspose.Cells TxtSaveOptions CSV conversion | C# export workbook to CSV ignoring comments | Aspose.Cells ExportCellComments property | .NET save Excel as CSV Aspose.Cells | CSV output without cell comments Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an XLSX workbook, configures TxtSaveOptions for CSV format (defaulting ExportCellComments to false), and saves the file as CSV, ensuring that cell comments are omitted from the output.
    class WorkbookToCsv
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.csv";

                // Ensure the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure CSV save options (using TxtSaveOptions for CSV format)
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv);

                // Save the workbook as CSV
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook successfully saved as CSV to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
