// Title: Load a macro‑enabled XLSM workbook and extract its VBA project using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load XLSM file and get VBA project | retrieve VBA modules from a macro-enabled workbook with Aspose.Cells | how to check if an XLSM workbook contains a VBA project in C# | list names of VBA modules in an Excel .xlsm using Aspose.Cells | C# code to access workbook.VbaProject property in Aspose.Cells | exception handling when loading VBA project from XLSM with Aspose.Cells
// Tags: load macro-enabled workbook Aspose.Cells C# | access VbaProject property Aspose.Cells | enumerate VBA modules Aspose.Cells | check VBA project existence XLSM Aspose.Cells | handle workbook loading errors Aspose.Cells | read VBA module names from XLSM using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VbaProject and VbaModule types

// The example verifies the presence of a .xlsm file, loads it into an Aspose.Cells Workbook, obtains the associated VbaProject, confirms its existence, and then outputs the count of VBA modules along with each module's name, while gracefully handling any runtime exceptions.
class Program
{
    static void Main()
    {
        string filePath = "YourWorkbook.xlsm";

        // Ensure the workbook file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the XLSM workbook from disk.
            Workbook workbook = new Workbook(filePath);

            // Obtain the VBA project associated with the workbook.
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Display the number of VBA modules in the project.
            Console.WriteLine($"Number of VBA modules: {vbaProject.Modules.Count}");

            // Iterate through modules and output their names.
            foreach (VbaModule module in vbaProject.Modules)
            {
                Console.WriteLine($"Module Name: {module.Name}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
