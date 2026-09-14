// Title: Load an existing Excel (.xlsx) file into an Aspose.Cells Workbook in C# with automatic fallback to a new workbook
// AI Prompts: Generate C# code that opens a .xlsx file using the Aspose.Cells Workbook constructor, checks if the file exists, and creates a new Workbook when the file is missing. | Add robust try‑catch handling around the workbook initialization and print the name of the first worksheet to the console. | Refactor the example to accept the Excel file path as a command‑line argument and display a clear message if the file cannot be found.
// Common Searches: aspocells c# load existing workbook from file path with fallback if file not found | how to open an .xlsx file using Aspose.Cells and handle missing file in .NET | c# Aspose.Cells create new workbook when specified Excel file does not exist | example of checking file existence before loading workbook with Aspose.Cells | aspocells exception handling when loading workbook from disk
// Tags: Aspose.Cells load workbook from .xlsx file | fallback create new workbook when file missing | check file existence before Aspose.Cells Workbook initialization | exception handling Aspose.Cells workbook loading | access first worksheet Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample checks whether a given .xlsx file exists; if it does, it loads the file into an Aspose.Cells Workbook, otherwise it creates a new Workbook. It then accesses the first worksheet, prints its name, and wraps the process in try‑catch error handling.
class Program
{
    static void Main()
    {
        // Path to the existing Excel file
        string filePath = @"C:\Path\To\YourFile.xlsx";

        try
        {
            Workbook workbook;

            // Load the workbook if the file exists; otherwise create a new one
            if (File.Exists(filePath))
            {
                workbook = new Workbook(filePath);
                Console.WriteLine($"Workbook loaded from '{filePath}'.");
            }
            else
            {
                workbook = new Workbook();
                Console.WriteLine($"File not found: '{filePath}'. A new workbook has been created.");
            }

            // The workbook object is now ready for further processing
            // Example: access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"First worksheet name: {sheet.Name}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors (e.g., Aspose.Cells specific exceptions)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
