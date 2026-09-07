// Title: How to list COM references in an Excel VBA project using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, accesses its VbaProject, and prints each COM reference name. | Show how to safely loop through the VbaProject.References collection using dynamic typing and handle errors in a .NET console application.
// Common Searches: c# aspocells list vba project com libraries | how to read VBA reference names from an Excel file using Aspose.Cells | enumerate VBA references in a workbook with .NET | retrieve COM reference list from Excel VBA project in C#
// Tags: Aspose.Cells VbaProject reference extraction | C# read VBA COM library names | dynamic VBA reference handling C# | Excel VBA project reference list .NET | read VBA library names with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads an Excel workbook with Aspose.Cells, checks for an embedded VBA project, and iterates through its VbaProject.References collection. Using dynamic typing, it prints each reference's Name while gracefully handling any errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project contained in the workbook
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // Iterate through all references in the VBA project
                foreach (object reference in vbaProject.References)
                {
                    try
                    {
                        // Use dynamic to access the Name property without needing the exact type at compile time
                        dynamic refDyn = reference;
                        Console.WriteLine(refDyn.Name);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to read reference: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }

            // If you need to save any changes, uncomment the line below
            // workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
