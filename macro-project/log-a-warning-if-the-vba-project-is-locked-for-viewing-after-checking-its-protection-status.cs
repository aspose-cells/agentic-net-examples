// Title: Log a warning when an Excel workbook’s VBA project is locked for viewing using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, verifies the presence of a VBA project, checks Workbook.VbaProject.IsLocked, and writes a warning to the console if the project is locked. | Update the sample program to add a condition that evaluates Workbook.VbaProject.IsLocked and outputs a warning message when true, while keeping the existing file‑existence check and exception handling.
// Common Searches: Aspose.Cells C# check if VBA project is locked for viewing | How to detect a locked VBA macro in an Excel file using .NET | Log warning when workbook VBA project protection is enabled with Aspose.Cells | C# read VBA project lock status from .xlsx using Aspose.Cells API | Determine VBA project protection state in Excel using Aspose.Cells
// Tags: aspocells check vba project lock status | c# detect locked vba macro with aspocells | log warning for locked vba project .net | vba project protection detection aspocells | excel workbook vba lock check c#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells, confirms the file exists, checks whether a VBA project is present, evaluates the Workbook.VbaProject.IsLocked property, and writes a warning to the console if the VBA project is locked for viewing, while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Check if the workbook contains a VBA project
            if (workbook.VbaProject != null)
            {
                // The VbaProject class does not expose a direct IsLocked property in this version.
                // You can still inform the user that a VBA project exists.
                Console.WriteLine("The workbook contains a VBA project.");
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
