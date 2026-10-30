// Title: Copy only cell values from a source range to another worksheet using PasteOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy range values without formulas | How to use PasteOptions to paste only values in Aspose.Cells | Copy cells from one workbook to another ignoring formatting with Aspose.Cells | C# example copying A1:C10 values to a new workbook using Aspose.Cells | PasteType.Values usage in Aspose.Cells for .NET
// Tags: PasteOptions values copy range Aspose.Cells | copy cell values without formulas Aspose.Cells C# | range to range value-only paste Aspose.Cells | Aspose.Cells copy values only .NET | ignore formatting during copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a source workbook, creates an empty destination workbook, defines a source range (A1:C10) and a destination range starting at A1, configures PasteOptions with PasteType.Values to copy only the cell values (excluding formulas and formatting), and saves the result to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string sourcePath = "source.xlsx";
                string resultPath = "result.xlsx";

                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file '{sourcePath}' not found.");
                    return;
                }

                // Load the source workbook
                Workbook srcWorkbook = new Workbook(sourcePath);

                // Create a new (empty) destination workbook
                Workbook destWorkbook = new Workbook();

                // Get the first worksheet from each workbook
                Worksheet srcSheet = srcWorkbook.Worksheets[0];
                Worksheet destSheet = destWorkbook.Worksheets[0];

                // Define the source range (A1:C10) and destination range (starting at A1)
                Aspose.Cells.Range srcRange = srcSheet.Cells.CreateRange("A1:C10");
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1");

                // Configure paste options to copy only cell values
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.Values
                };

                // Perform the copy operation using the configured options
                srcRange.Copy(destRange, pasteOptions);

                // Ensure the directory for the result file exists
                string resultDir = Path.GetDirectoryName(resultPath);
                if (!string.IsNullOrEmpty(resultDir) && !Directory.Exists(resultDir))
                {
                    Directory.CreateDirectory(resultDir);
                }

                // Save the destination workbook
                destWorkbook.Save(resultPath);
                Console.WriteLine($"Data copied successfully to '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
