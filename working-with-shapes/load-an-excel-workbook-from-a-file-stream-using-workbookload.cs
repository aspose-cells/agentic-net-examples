// Title: How to load an Excel workbook from a file path using Aspose.Cells Workbook class in C#
// Common Searches: C# code to open an .xlsx file with Aspose.Cells without errors | How to safely instantiate Aspose.Cells Workbook for an existing Excel file | Example of reading worksheet names after loading Excel using Aspose.Cells | Prevent FileNotFoundException when loading Excel with Aspose.Cells in .NET | Console app loading Excel workbook via Aspose.Cells in C#
// Tags: Aspose.Cells load workbook from file path | C# Aspose.Cells open .xlsx file | Workbook constructor file existence validation | retrieve first worksheet name Aspose.Cells | exception handling Aspose.Cells workbook loading

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the specified .xlsx file exists, then creates a new Aspose.Cells Workbook instance with the file path to load the workbook. It demonstrates how to access the first worksheet's name and includes basic exception handling for robust file loading.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook directly from the file path (Aspose.Cells supports this)
            Workbook workbook = new Workbook(filePath);

            // At this point the workbook is loaded and can be manipulated
            // Example: output the name of the first worksheet
            if (workbook.Worksheets.Count > 0)
            {
                Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
            }
            else
            {
                Console.WriteLine("The workbook contains no worksheets.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine("An error occurred while processing the Excel file:");
            Console.WriteLine(ex.Message);
        }
    }
}
