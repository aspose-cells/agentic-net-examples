// Title: Load an Excel workbook from a file path and verify successful opening with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load workbook from local file path and check if it opened | how to catch exceptions when opening an Excel file with Aspose.Cells in .NET | verify number of worksheets after loading an .xlsx using Aspose.Cells C# | sample code for initializing Aspose.Cells Workbook with a file path | Aspose.Cells .NET example to confirm workbook load success
// Tags: load workbook from file path Aspose.Cells C# | exception handling workbook load Aspose.Cells | check worksheet count after workbook initialization Aspose.Cells | initialize Aspose.Cells Workbook with .xlsx path | verify workbook opened successfully Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates how to create a Workbook object by providing a file path to an .xlsx file, prints a success message and the worksheet count when loading succeeds, and catches any exceptions to report a failure.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = @"C:\Path\To\Your\Workbook.xlsx";

            try
            {
                // Load the workbook from the specified file path
                Workbook workbook = new Workbook(filePath);

                // If no exception is thrown, the workbook opened successfully
                Console.WriteLine("Workbook loaded successfully.");
                
                // Optional: you can perform additional checks, e.g., verify sheet count
                Console.WriteLine($"Number of worksheets: {workbook.Worksheets.Count}");
            }
            catch (Exception ex)
            {
                // If an exception occurs, the workbook failed to open
                Console.WriteLine($"Failed to load workbook. Error: {ex.Message}");
            }
        }
    }
}
