// Title: Load an XLSX workbook from a file path and confirm it contains worksheets using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, checks that the Workbook object is instantiated and has at least one worksheet, then prints the worksheet count. | Write a C# try‑catch example that loads a workbook from a given path using Aspose.Cells, validates successful loading, and logs any exception messages.
// Common Searches: asp.net load excel file with Aspose.Cells and verify worksheets exist | c# check if workbook loaded from path contains any sheets using Aspose.Cells | how to handle exceptions when opening an .xlsx file with Aspose.Cells in .NET | determine number of worksheets after initializing Aspose.Cells workbook from file | sample code for validating workbook initialization in Aspose.Cells C#
// Tags: load xlsx workbook Aspose.Cells C# | validate worksheet presence Aspose.Cells | exception handling Aspose.Cells file load | retrieve worksheet count Aspose.Cells | initialize workbook from file path .NET

using System;
using Aspose.Cells;

// Shows how to load an XLSX file into an Aspose.Cells Workbook in C#, verify that the workbook is instantiated and contains at least one worksheet, output the sheet count, and handle any loading errors.
class Program
{
    static void Main()
    {
        // Path to the XLSX file to be loaded
        string filePath = @"C:\Path\To\Your\File.xlsx";

        try
        {
            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Verify successful initialization
            if (workbook != null && workbook.Worksheets.Count > 0)
            {
                Console.WriteLine("Workbook loaded successfully.");
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            }
            else
            {
                Console.WriteLine("Workbook loaded, but it contains no worksheets.");
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
