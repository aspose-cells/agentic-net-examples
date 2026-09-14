// Title: Load an Excel workbook from a file path and verify it opened successfully using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a Workbook from a given .xlsx file path and prints the total number of worksheets. | Show how to wrap the Workbook constructor in a try‑catch block to capture and log any loading errors with Aspose.Cells. | Provide a snippet that confirms a workbook was loaded by checking the Worksheets.Count property after using the string‑based constructor.
// Common Searches: aspnet load excel file with Aspose.Cells and get worksheet count | how to catch errors when opening a workbook using Aspose.Cells C# | verify that an .xlsx file was loaded successfully with Aspose.Cells .NET | example code for loading a workbook from a path and handling exceptions Aspose.Cells
// Tags: load workbook Aspose.Cells | verify workbook opened Aspose.Cells | handle workbook constructor exception Aspose.Cells | retrieve worksheets count Aspose.Cells | C# open .xlsx with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    // // Demonstrates loading "example.xlsx" into an Aspose.Cells Workbook, outputting the worksheet count, and handling any exceptions that occur during the load.
    class Program
    {
        static void Main()
        {
            // Path to the Excel file to be loaded
            string filePath = "example.xlsx";

            try
            {
                // Load the workbook using the string constructor (loads the file)
                Workbook workbook = new Workbook(filePath);

                // Verify successful load by checking the number of worksheets
                Console.WriteLine($"Workbook loaded successfully. Worksheets count: {workbook.Worksheets.Count}");
            }
            catch (Exception ex)
            {
                // Report any errors that occurred during loading
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
            }
        }
    }
}
