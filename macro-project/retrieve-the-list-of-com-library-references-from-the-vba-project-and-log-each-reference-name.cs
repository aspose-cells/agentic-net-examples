// Title: C# example to list and log COM library references from an Excel VBA project using Aspose.Cells
// Common Searches: C# code to list VBA references in an Excel workbook with Aspose.Cells | how to retrieve COM library reference names from an Excel file using Aspose.Cells | sample for enumerating VBA project references in .NET | Aspose.Cells example to read VBA reference collection in C# | extract VBA reference names from an Excel workbook programmatically
// Tags: Aspose.Cells read VBA references C# | enumerate COM references Excel Aspose | list VbaProject.References .NET | extract VBA reference names C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

// The program verifies the input Excel file, loads it with Aspose.Cells, accesses its VbaProject, and iterates through the VbaProject.References collection, printing each reference's Name property while handling missing files, absent VBA projects, and runtime exceptions.
class Program
{
    static void Main()
    {
        // Path to the input workbook
        string workbookPath = "input.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"File not found: {workbookPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project (may be null if no VBA is present)
            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject == null)
            {
                Console.WriteLine("No VBA project found in the workbook.");
                return;
            }

            // Iterate through VBA references
            foreach (var reference in vbaProject.References)
            {
                // The Name property contains the reference name (e.g., "Excel", "Outlook")
                Console.WriteLine($"Reference Name: {reference.Name}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
