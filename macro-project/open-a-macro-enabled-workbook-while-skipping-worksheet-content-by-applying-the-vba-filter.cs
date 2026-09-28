// Title: Open a macro-enabled .xlsm workbook in C# with Aspose.Cells, loading only the VBA project and skipping worksheet cells
// AI Prompts: Use Aspose.Cells LoadOptions in C# to open an .xlsm file with the VBA filter enabled, then enumerate the VBA modules without loading any worksheet data. | Create a C# program that checks for an .xlsm file, loads only its VBA project via Aspose.Cells, and prints each module name while ignoring sheet content.
// Common Searches: Aspose.Cells C# load only VBA project from .xlsm file | skip loading worksheets when opening macro-enabled Excel workbook with Aspose.Cells | retrieve VBA module names without reading cell data using Aspose.Cells | how to use LoadOptions VBA filter in Aspose.Cells C# example | open .xlsm in C# and list VBA modules without loading sheets
// Tags: Aspose.Cells LoadOptions VBA filter | C# load only VBA project .xlsm | skip worksheet data Aspose.Cells | enumerate VBA modules Aspose.Cells | macro-enabled workbook loading C#

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that a macro-enabled .xlsm file exists, loads it with Aspose.Cells using the VBA filter to avoid loading worksheet cells, accesses the VbaProject, iterates through its modules, and prints each module name while handling possible errors.
class Program
{
    static void Main()
    {
        const string filePath = "MacroWorkbook.xlsm";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load the workbook (Aspose.Cells auto‑detects the format)
            workbook = new Workbook(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        try
        {
            // Access the VBA project (if any) and list its modules
            var vbaProject = workbook.VbaProject;
            if (vbaProject != null && vbaProject.Modules != null)
            {
                foreach (var module in vbaProject.Modules)
                {
                    Console.WriteLine($"Module: {module.Name}");
                }
            }
            else
            {
                Console.WriteLine("No VBA project or modules found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing VBA project: {ex.Message}");
        }
    }
}
