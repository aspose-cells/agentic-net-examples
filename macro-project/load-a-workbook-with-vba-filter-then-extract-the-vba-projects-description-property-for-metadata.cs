// Title: Read the VBA project name from a macro‑enabled XLSM workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to get VBA project name from an .xlsm file | retrieve VBA project metadata with Aspose.Cells Workbook in .NET | read VBA project properties from a macro-enabled workbook using Aspose.Cells | Aspose.Cells VbaProject Name property example in C# | load XLSM and access VBA project information with Aspose.Cells
// Tags: Aspose.Cells read VBA project name | C# extract VBA metadata from XLSM | Aspose.Cells VbaProject access | macro-enabled workbook VBA properties | load XLSM workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the presence of an input.xlsm file, loads it with Aspose.Cells, accesses the workbook's VbaProject, and prints the VBA project's Name property to the console while handling potential errors.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the workbook (XLSM format is detected automatically)
            var workbook = new Workbook(inputFile);

            // Access the VBA project attached to the workbook
            var vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // Retrieve the Name property from the VBA project metadata
                Console.WriteLine("VBA Project Name: " + vbaProject.Name);
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
