// Title: Load an XLSM workbook and read the VBA project name with Aspose.Cells for .NET (C#)
// Common Searches: C# how to read VBA project name from an .xlsm file using Aspose.Cells | Aspose.Cells retrieve VBA project information from a macro-enabled workbook in .NET | example code to get VBA project name from Excel macro file with Aspose.Cells C# | check if VBA project exists and obtain its name using Aspose.Cells library | load macro-enabled workbook and access VbaProject.Name property in C#
// Tags: Aspose.Cells load macro-enabled workbook | C# read VBA project name Aspose.Cells | VbaProject.Name property extraction | handle missing VBA project Aspose.Cells | extract VBA metadata from XLSM using .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The sample verifies the XLSM file exists, loads it with Aspose.Cells, checks for a VBA project, and prints the VbaProject.Name value or a message when no VBA project is present.
class Program
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook (VBA project is loaded automatically if present)
            Workbook workbook = new Workbook(inputPath);

            // Extract the VBA project's name (if a VBA project exists)
            string vbaInfo = workbook.VbaProject != null
                ? $"VBA Project Name: {workbook.VbaProject.Name}"
                : "No VBA project found";

            Console.WriteLine(vbaInfo);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
