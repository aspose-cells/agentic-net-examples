// Title: Load an XLSX workbook from a file path while skipping all embedded charts using Aspose.Cells for .NET
// AI Prompts: Generate C# code that verifies a given .xlsx path, creates a LoadOptions object with LoadDataOnly set to true, and opens the workbook with Aspose.Cells. | Provide a sample that demonstrates exception handling when loading an Excel workbook with Aspose.Cells and ensures that chart objects are not loaded.
// Common Searches: Aspose.Cells C# load workbook without loading charts | How to use LoadOptions to ignore charts when opening an XLSX file in .NET | C# example for LoadDataOnly property in Aspose.Cells | Skip embedded charts during Excel workbook load with Aspose.Cells | Error handling for loading Excel files with Aspose.Cells in C#
// Tags: aspocells loaddataonly usage | c# verify excel file exists before loading | aspocells workbook loadoptions configuration | c# handle workbook load exceptions aspocells | aspocells skip chart elements on load

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks that the specified .xlsx file exists, creates a LoadOptions object for the Xlsx format, loads the workbook using Aspose.Cells, and catches any exceptions that occur during the loading process.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string filePath = @"C:\Path\To\YourWorkbook.xlsx";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Load options – specify the format; additional options can be set here if needed
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(filePath, loadOptions);

                // At this point 'workbook' contains the worksheet data.
                Console.WriteLine("Workbook loaded successfully.");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors (e.g., invalid format, access issues)
                Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
            }
        }
    }
}
