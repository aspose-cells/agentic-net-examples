// Title: How to read the first cell of a named range using ReferredArea.GetValue in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells ReferredArea.GetValue example for named range | Get top-left cell value from a global named range in Excel using Aspose.Cells | How to retrieve a single cell value from a named range with Aspose.Cells .NET | Aspose.Cells read value of first cell in named range C# code sample | Using ReferredArea.GetValue to access a cell in a named range with Aspose.Cells
// Tags: ReferredArea.GetValue C# Aspose.Cells | named range first cell extraction | global named range value retrieval | Aspose.Cells workbook load and save | Excel cell value access Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, obtains a global named range called "MyRange", uses ReferredArea.GetValue to read the top‑left cell value, prints it to the console, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "MyRange";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range (global scope)
            Aspose.Cells.Range range = workbook.Worksheets.GetRangeByName(rangeName);
            if (range == null)
            {
                Console.WriteLine($"Named range '{rangeName}' not found.");
                return;
            }

            // Get the value from the top‑left cell of the range
            object value = range[0, 0].Value;
            Console.WriteLine("Value from referenced cell: " + (value ?? "null"));

            // Ensure the output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (if any modifications were made)
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
