// Title: How to limit Aspose.Cells smart marker processing to a named range using WorkbookDesigner.SetRange in C#
// Common Searches: Aspose.Cells C# set smart marker processing range with WorkbookDesigner | limit smart markers to a named range using SetRange method | example of WorkbookDesigner.SetRange for Excel smart markers | restrict Aspose.Cells smart marker scope to specific cells in C# | parameters for SetRange when processing smart markers in Aspose.Cells
// Tags: limit smart marker processing | SetRange API usage | named range smart markers | Aspose.Cells C# range restriction | WorkbookDesigner smart marker scope

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This example loads an Excel workbook, creates a WorkbookDesigner, optionally limits smart marker processing to a named range with SetRange, processes the markers, and saves the result, handling missing files and runtime errors.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Create a WorkbookDesigner for smart marker processing
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // If you need to limit processing to a named range, uncomment and adjust the following line:
                // designer.SetRange("DataRange", 0, 0, 0, 0); // parameters: name, startRow, startColumn, totalRows, totalColumns

                // Process the smart markers
                designer.Process();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
