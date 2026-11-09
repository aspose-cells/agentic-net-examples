// Title: How to retrieve a VBA project from a loaded Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get VBA project from existing workbook | C# check if Excel file contains VBA macros with Aspose.Cells | Retrieve VBA modules count from workbook using Aspose.Cells API | Load .xlsx and access VbaProject property in Aspose.Cells for .NET | Handle missing VBA project when reading Excel with Aspose.Cells C#
// Tags: Aspose.Cells VbaProject extraction C# | detect VBA macros in Excel workbook using Aspose.Cells | count VBA modules via Aspose.Cells VbaProject | load workbook and read VbaProject property | exception handling for missing VBA project Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example checks for the presence of 'input.xlsx', loads it into an Aspose.Cells Workbook, accesses the Workbook.VbaProject property, reports the number of VBA modules if a project exists, and gracefully handles cases where no VBA project is found or other errors occur.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: The file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Retrieve the VBA project from the workbook (may be null)
            VbaProject vbaProject = workbook.VbaProject;

            // Check if a VBA project exists and output module count
            if (vbaProject != null)
            {
                Console.WriteLine($"VBA project found. Module count: {vbaProject.Modules.Count}");
            }
            else
            {
                Console.WriteLine("No VBA project present in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
