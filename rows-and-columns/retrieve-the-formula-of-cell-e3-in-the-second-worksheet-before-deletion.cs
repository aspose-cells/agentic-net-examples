// Title: Get the formula of cell E3 from the second worksheet using Aspose.Cells for .NET before deleting the sheet
// Common Searches: Aspose.Cells C# read formula from cell E3 on second worksheet before removing sheet | how to obtain a cell formula in a specific worksheet with Aspose.Cells .NET | retrieve Excel formula from E3 in sheet index 1 using Aspose.Cells | C# Aspose.Cells get cell formula then delete worksheet | sample code to read formula from a hidden sheet with Aspose.Cells
// Tags: Aspose.Cells extract cell formula .NET | access second worksheet Aspose.Cells | remove worksheet after data extraction Aspose.Cells | C# get Excel formula before sheet deletion | Aspose.Cells workbook cell formula retrieval

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, checks for a second worksheet, reads the formula in cell E3 of that sheet, prints it, then removes the second worksheet and saves the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a sample if missing
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a sample workbook.");
                    var sampleWb = new Workbook();
                    // Add a second worksheet
                    sampleWb.Worksheets.Add();
                    var secondSheet = sampleWb.Worksheets[1];
                    // Place a sample formula in E3
                    secondSheet.Cells["E3"].Formula = "=SUM(A1:A5)";
                    sampleWb.Save(inputPath);
                }

                // Load the workbook from the input file
                var workbook = new Workbook(inputPath);

                // Verify that a second worksheet exists
                if (workbook.Worksheets.Count > 1)
                {
                    var secondSheet = workbook.Worksheets[1];

                    // Retrieve the formula from cell E3 before any modification
                    string e3Formula = secondSheet.Cells["E3"].Formula;
                    Console.WriteLine($"Formula in E3 before deletion: {e3Formula}");

                    // Delete the second worksheet as required
                    workbook.Worksheets.RemoveAt(1);
                }
                else
                {
                    Console.WriteLine("The workbook does not contain a second worksheet.");
                }

                // Save the modified workbook to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Modified workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
