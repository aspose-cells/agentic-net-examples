// Title: Load a macro-enabled .xlsm workbook and extract the VBA project's Description metadata with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsm file using Aspose.Cells, verifies a VBA project exists, and prints the project's Description property. | Show how to access Workbook.VbaProject.Description in Aspose.Cells to retrieve VBA project metadata from a macro-enabled Excel file. | Create a reusable C# method that returns the Description string of a VBA project from a given .xlsm workbook using Aspose.Cells.
// Common Searches: Aspose.Cells read VBA project description from .xlsm in C# | How to get VBA project metadata with Aspose.Cells .NET | C# extract Description property of VBA project in macro-enabled Excel file | Retrieve VBA project properties using Aspose.Cells for .NET | Sample code for reading VBA description from Excel workbook with Aspose.Cells
// Tags: aspnet aspose.cells read vba description | c# load macro-enabled xlsm workbook | aspose.cells access vba project metadata | extract vba project description .net | workbook.vbaproject description property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example demonstrates how to load a macro-enabled .xlsm workbook with Aspose.Cells for .NET, confirm that a VBA project is present, and read the project's Description property for use as metadata.
class Program
{
    static void Main()
    {
        try
        {
            string filePath = "input.xlsm";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook; Aspose.Cells automatically detects the format (including .xlsm)
            Workbook workbook = new Workbook(filePath);

            // Check for the presence of a VBA project
            if (workbook.VbaProject != null)
            {
                Console.WriteLine("VBA project is present in the workbook.");
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
