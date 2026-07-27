// Title: Load a .xls (Excel 97‑2003) workbook in Aspose.Cells for .NET using LoadOptions
// Description: Shows how to check for a .xls file, create LoadOptions with LoadFormat.Excel97To2003, load the workbook, print the first worksheet name, and gracefully handle exceptions.
// Keywords: Aspose.Cells | LoadOptions | Excel97To2003 | .xls | C# | legacy Excel | LoadFormat | open workbook | file existence check
// Common Searches: Aspose.Cells load .xls file C# | LoadOptions Excel97To2003 example | How to open Excel 97‑2003 workbook with Aspose.Cells | Check file exists before loading Aspose.Cells | Specify workbook format in Aspose.Cells .NET
// Developer Intent: Open a legacy Excel 97‑2003 workbook by explicitly setting the format in LoadOptions.
// Use Cases: Read data from an old .xls file in a .NET application. | Validate the presence of a legacy workbook before processing. | Handle format‑specific loading errors when working with .xls files.
// AI Prompts: Write C# code that loads an Excel 97‑2003 (.xls) workbook with Aspose.Cells using LoadOptions and prints all worksheet names. | Provide a robust Aspose.Cells snippet that checks file existence, loads a .xls file with the correct LoadFormat, and logs any exceptions. | Explain how to use LoadOptions (or its constructor) to set LoadFormat.Excel97To2003 when opening a workbook in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadExample
{
    // Shows how to check for a .xls file, create LoadOptions with LoadFormat.Excel97To2003, load the workbook, print the first worksheet name, and gracefully handle exceptions.
    class Program
    {
        static void Main()
        {
            // Path to the Excel 97‑2003 workbook to be loaded
            string filePath = "sample.xls";

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            try
            {
                // Create LoadOptions specifying the Excel 97‑2003 format via constructor
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

                // Load the workbook using the file path and the configured LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Example usage: display the name of the first worksheet
                Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine("An error occurred while loading the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
