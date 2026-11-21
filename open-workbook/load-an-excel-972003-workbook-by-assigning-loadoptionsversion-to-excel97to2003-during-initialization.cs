// Title: Load a legacy Excel 97‑2003 (.xls) workbook and convert it to XLSX with Aspose.Cells for .NET (C#)
// Common Searches: how to open an .xls file with Aspose.Cells in C# and save as .xlsx | Aspose.Cells C# example for loading a legacy Excel 97-2003 workbook | convert Excel 97-2003 workbook to XLSX using Aspose.Cells .NET | C# code to read .xls file and export to .xlsx with Aspose.Cells | Aspose.Cells load xls format and handle file not found exception in C#
// Tags: load legacy xls workbook Aspose.Cells | convert xls to xlsx C# | Aspose.Cells workbook loading options | save workbook as xlsx Aspose.Cells | C# read Excel 97-2003 file Aspose.Cells | handle missing input file Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that 'input.xls' exists, loads the legacy .xls workbook with Aspose.Cells, and then saves it as 'output.xlsx' in XLSX format, handling any loading or saving errors.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xls";
        const string outputFile = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook; Aspose.Cells automatically detects the format
            workbook = new Workbook(inputFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Save the workbook in XLSX format
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook successfully saved as '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
