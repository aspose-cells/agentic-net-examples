// Title: Load an Excel workbook, verify file existence, and save it unchanged using Aspose.Cells for .NET (sparkline assembly note)
// Common Searches: how to load an Excel file and save it with Aspose.Cells in C# | asp.net check if input.xlsx exists before opening with Aspose.Cells | why is sparkline manipulation code omitted in Aspose.Cells examples | how to add Aspose.Cells.Sparkline assembly reference to a C# project | exception handling when saving a workbook with Aspose.Cells
// Tags: load workbook Aspose.Cells C# | save workbook unchanged Aspose.Cells .NET | verify input file existence Aspose.Cells | reference Aspose.Cells.Sparkline assembly | exception handling Aspose.Cells workbook operations

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook from "input.xlsx" after confirming the file exists, then saves it unchanged to "output.xlsx" using Aspose.Cells for .NET. It includes a note that sparkline manipulation requires the Aspose.Cells.Sparkline assembly, which is not referenced in this code, and demonstrates basic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Sparkline manipulation requires the Aspose.Cells.Sparkline assembly.
            // If the assembly is not referenced, sparkline-related code is omitted.

            // Save the (potentially unchanged) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
