// Title: Load an Excel workbook from a file path with Aspose.Cells for .NET (C#)
// Description: Shows how to create a Workbook from a file path, verify the load by reading the worksheet count, and catch any exceptions in C# using Aspose.Cells.
// Keywords: Aspose.Cells | C# | load workbook | file path | verify workbook opened | exception handling | worksheet count | Excel .xlsx | Aspose.Cells example
// Common Searches: Aspose.Cells load workbook C# | C# read Excel file with Aspose.Cells | check if Excel file opened Aspose.Cells | Aspose.Cells get worksheet count | handle file not found Aspose.Cells
// Developer Intent: Open an Excel file into a Workbook object and confirm it loaded successfully by accessing its properties.
// Use Cases: Load a user‑provided .xlsx file and display the number of worksheets before further processing. | Validate that a template workbook exists and is readable before populating it with data. | Log detailed errors when the file path is invalid, missing, or the workbook is corrupted.
// AI Prompts: Generate C# code that loads an Excel workbook from a given path using Aspose.Cells and prints each worksheet name. | Create a reusable method that attempts to load a workbook, returns a boolean success flag, and logs any exceptions. | Show how to load a password‑protected Excel file with Aspose.Cells and verify that the load succeeded.

using System;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    // Shows how to create a Workbook from a file path, verify the load by reading the worksheet count, and catch any exceptions in C# using Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "example.xlsx";

            try
            {
                // Load the workbook using the string constructor (loads the file)
                Workbook workbook = new Workbook(filePath);

                // Verify that the workbook opened by accessing a property
                Console.WriteLine("Workbook loaded successfully.");
                Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);
            }
            catch (Exception ex)
            {
                // If loading fails, output the error message
                Console.WriteLine("Failed to load workbook: " + ex.Message);
            }
        }
    }
}
