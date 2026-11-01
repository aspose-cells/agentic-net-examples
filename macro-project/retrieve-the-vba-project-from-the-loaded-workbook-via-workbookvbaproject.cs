// Title: How to access and list VBA modules from an Excel workbook using Workbook.VbaProject in C# (Aspose.Cells)
// Common Searches: C# Aspose.Cells get VBA project from Excel workbook | list VBA modules in .xlsx file using Aspose.Cells .NET | Workbook.VbaProject returns null why Aspose.Cells | check if Excel file contains VBA project with Aspose.Cells C# | extract VBA module names from workbook using Aspose.Cells API
// Tags: access workbook VbaProject Aspose.Cells | enumerate VBA modules C# Aspose.Cells | null VbaProject handling Aspose.Cells | extract VBA project from Excel .NET | detect missing VBA macros Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an Excel workbook with Aspose.Cells, accesses its VbaProject via the Workbook.VbaProject property, verifies the project exists, iterates through each VBA module to print its name, and includes error handling for missing projects or runtime exceptions.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Retrieve the VBA project (may be null if none exists)
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Iterate through VBA modules and output their names
            foreach (var module in vbaProject.Modules)
            {
                Console.WriteLine($"Module name: {module.Name}");
            }
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
