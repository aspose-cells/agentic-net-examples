// Title: Process smart markers on every worksheet of an Excel workbook with Aspose.Cells WorkbookDesigner in C#
// Common Searches: C# Aspose.Cells process smart markers on all worksheets | How to run WorkbookDesigner.Process for each sheet in an Excel file using Aspose.Cells | Iterate Excel worksheets and apply smart marker processing with Aspose.Cells .NET | Example of batch smart marker replacement across multiple sheets in C#
// Tags: WorkbookDesigner.Process multiple worksheets | smart marker batch processing Aspose.Cells | C# iterate Excel sheets for smart markers | Aspose.Cells process all smart markers

using System;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerProcessing
{
    // The sample loads an Excel file, creates a WorkbookDesigner, invokes Process() to replace all smart markers across every worksheet, and saves the result, with basic file‑existence checks and exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook containing smart markers
                Workbook workbook = new Workbook(inputPath);

                // Create a WorkbookDesigner for the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Process all smart markers in the workbook
                designer.Process();

                // Save the processed workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
