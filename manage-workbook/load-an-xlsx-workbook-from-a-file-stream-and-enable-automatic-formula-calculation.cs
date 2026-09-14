// Title: Load an XLSX workbook from a read‑only FileStream and calculate all formulas with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an XLSX file via a read‑only FileStream, loads it into an Aspose.Cells Workbook, and invokes CalculateFormula to evaluate every formula. | Show how to verify that a workbook file path exists, handle missing‑file errors, and implement exception handling when loading a workbook from a stream using Aspose.Cells.
// Common Searches: C# Aspose.Cells load workbook from FileStream and calculate formulas | how to enable automatic formula calculation when opening an XLSX with Aspose.Cells | Aspose.Cells example for reading an XLSX file as a read‑only stream | validate file path before loading workbook using Aspose.Cells .NET | exception handling pattern for workbook loading from stream Aspose.Cells
// Tags: open xlsx file as read‑only stream Aspose.Cells | calculate workbook formulas Aspose.Cells | file path existence check Aspose.Cells .NET | exception handling for workbook loading Aspose.Cells | stream‑based workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example validates a non‑empty, existing file path, opens the XLSX file with a read‑only FileStream, loads it into an Aspose.Cells Workbook, forces immediate formula evaluation via CalculateFormula, outputs a success message, and catches any processing exceptions.
    public class WorkbookProcessor
    {
        /// <param name="filePath">Full path to the XLSX file.</param>
        public void LoadWorkbookAndEnableAutoCalc(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("Error: File path is null or empty.");
                return;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found at '{filePath}'.");
                return;
            }

            try
            {
                // Open the XLSX file as a read‑only stream
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    // Load the workbook from the stream
                    Workbook workbook = new Workbook(stream);

                    // Force an immediate calculation of all formulas
                    workbook.CalculateFormula();

                    Console.WriteLine("Workbook loaded and formulas calculated successfully.");
                    // Further processing can be done here
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Example usage: provide the path to an existing XLSX file.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            WorkbookProcessor processor = new WorkbookProcessor();
            processor.LoadWorkbookAndEnableAutoCalc(filePath);
        }
    }
}
