// Title: How to list COM library references from a VBA project in a macro-enabled Excel workbook using Aspose.Cells for .NET
// Common Searches: C# list VBA references in an .xlsm file with Aspose.Cells | Aspose.Cells enumerate COM libraries from VBA project | retrieve VBA project reference names using Aspose.Cells .NET | get list of VbaProject.References in a macro-enabled workbook C# | how to read VBA reference collection from Excel workbook using Aspose.Cells
// Tags: enumerate VBA references Aspose.Cells | list COM libraries in .xlsm workbook | Aspose.Cells VbaProject.References iteration | C# extract VBA reference names | macro-enabled workbook VBA reference extraction

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The example loads a macro‑enabled Excel file with Aspose.Cells, accesses its VbaProject, iterates through the VbaProject.References collection, and writes each COM library reference name to the console, handling missing VBA projects and runtime errors.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the macro-enabled workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project (may be null if no macros are present)
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject != null)
            {
                // Iterate through the VBA references and output their names
                foreach (var reference in vbaProject.References)
                {
                    Console.WriteLine(reference.Name);
                }
            }
            else
            {
                Console.WriteLine("No VBA project found in the workbook.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
