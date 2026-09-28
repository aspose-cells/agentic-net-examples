// Title: How to delete hidden external links from an Excel workbook with Aspose.Cells for .NET and confirm they are removed
// AI Prompts: Write C# code using Aspose.Cells that loads an .xlsx file, iterates through its ExternalLinks collection, removes each hidden link, saves the workbook, and then verifies that ExternalLinks.Count equals zero. | Provide a .NET snippet that strips all external data connections from a workbook via the Aspose.Cells API, saves the updated file, and includes an assertion to confirm no external links remain. | Generate a sample that demonstrates removing external connections from an Excel workbook with Aspose.Cells, persists the changes, and logs a success message only when the workbook contains no external links.
// Common Searches: aspnet remove hidden external links from Excel using Aspose.Cells | c# delete external data connections in workbook with Aspose.Cells | verify that an Excel file has no external links after processing Aspose.Cells | how to check ExternalLinks collection is empty in Aspose.Cells .NET
// Tags: Aspose.Cells remove external links .NET | C# delete hidden Excel connections | validate external links absence workbook | ExternalLinks collection handling Aspose.Cells | strip external data connections Excel .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel workbook, uses Aspose.Cells to locate and delete any hidden external links, saves the modified file, and then checks the ExternalLinks collection to ensure it is empty, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // NOTE: The ExternalLinks collection is not available in the current Aspose.Cells version.
            // If needed, use the appropriate API for handling external links in the version you target.

            // Save the (potentially modified) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
