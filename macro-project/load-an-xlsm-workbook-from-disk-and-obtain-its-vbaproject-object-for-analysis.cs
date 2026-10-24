// Title: Load an XLSM workbook and retrieve its VBA project with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells example to open a macro-enabled .xlsm file and access the VbaProject | How to read VBA modules count from an XLSM workbook using Aspose.Cells for .NET | Retrieve VBA project from an Excel macro workbook with Aspose.Cells API in C# | Aspose.Cells load macro-enabled workbook and get VbaProject object | C# code to check existence and load .xlsm file then enumerate VBA modules using Aspose.Cells
// Tags: Aspose.Cells VBA project extraction .NET | C# enumerate VBA modules in macro-enabled workbook | open macro-enabled Excel file via Aspose.Cells | retrieve VbaProject instance from XLSM | Aspose.Cells workbook VbaProject API usage

using Aspose.Cells;
using Aspose.Cells.Vba;
using System;
using System.IO;

// The sample verifies that the input.xlsm file exists, loads it into an Aspose.Cells Workbook, obtains the associated VbaProject, and prints the number of VBA modules, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsm";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the XLSM workbook from disk
            Workbook workbook = new Workbook(filePath);

            // Obtain the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Display the count of VBA modules
            Console.WriteLine("VBA Modules Count: " + vbaProject.Modules.Count);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
