// Title: How to disable automatic XML map refresh in Aspose.Cells for .NET while performing bulk cell updates
// Common Searches: disable xml map refresh in Aspose.Cells when writing thousands of cells c# | suspend automatic xml map refresh to speed up bulk updates in .NET Excel workbook | Aspose.Cells bulk cell write without triggering xml map refresh performance tip | c# code to turn off xml map refresh during large data import with Aspose.Cells | improve Excel processing time by disabling xml map refresh in Aspose.Cells
// Tags: disable xml map refresh aspose.cells | bulk cell update performance asp.net | suspend xml map refresh during data import | optimize workbook processing speed aspose.cells | excel xml map refresh control c# | large dataset write aspose.cells performance

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates loading or creating an Excel workbook with Aspose.Cells, turning off automatic XML map refresh, writing values to a large range of cells (10,000 rows × 50 columns) to illustrate bulk updates, and then saving the workbook, thereby showing how to improve processing speed by suspending XML map refresh.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "Input.xlsx";
            const string outputFile = "Output.xlsx";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found. Creating a new workbook.");
                }

                // Load existing workbook or create a new one if the file is missing
                Workbook workbook = File.Exists(inputFile) ? new Workbook(inputFile) : new Workbook();

                // Access the first worksheet and its cells collection
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Perform bulk updates on a large range of cells
                for (int row = 0; row < 10000; row++)
                {
                    for (int col = 0; col < 50; col++)
                    {
                        cells[row, col].PutValue(row * col);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
