// Title: Open an Excel 95 (.xls) file with Aspose.Cells for .NET using LoadOptions.Version
// Description: Shows how to load a legacy Excel 95 workbook in C# by creating a LoadOptions object, setting its Version to ExcelVersion.Excel95, and passing it to the Workbook constructor. The sample checks file existence, prints the first worksheet name, optionally saves the workbook, and includes full exception handling.
// Keywords: Aspose.Cells | C# | Excel95 | LoadOptions.Version | legacy .xls | open Excel 95 | read old Excel format | Workbook constructor | ExcelVersion.Excel95 | exception handling
// Common Searches: Aspose.Cells open Excel 95 file | LoadOptions.Version Excel95 C# | Read legacy .xls with Aspose.Cells | How to load old Excel format .NET | Open Excel 95 workbook using Aspose.Cells
// Developer Intent: Load a legacy Excel 95 workbook in C# and optionally save it after confirming successful load.
// Use Cases: Validate that a legacy .xls file exists before attempting to load it. | Load an Excel 95 workbook by setting LoadOptions.Version to ensure correct format detection. | Display the name of the first worksheet to verify the file was read correctly. | Save the loaded workbook to a new location to confirm successful import. | Handle FileNotFoundException and other runtime errors gracefully.
// AI Prompts: Write C# code that opens an Excel 95 (.xls) file with Aspose.Cells by setting LoadOptions.Version to ExcelVersion.Excel95 before creating the Workbook. | Show how to implement robust exception handling (FileNotFoundException, generic Exception) when loading a legacy Excel file using Aspose.Cells. | Explain steps to save a workbook loaded from Excel 95 to a different file path with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load a legacy Excel 95 workbook in C# by creating a LoadOptions object, setting its Version to ExcelVersion.Excel95, and passing it to the Workbook constructor. The sample checks file existence, prints the first worksheet name, optionally saves the workbook, and includes full exception handling.
class OpenExcel95Example
{
    static void Main()
    {
        // Path to the Excel 95 file (xls format)
        string inputFile = "SampleExcel95.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Input file not found: {inputFile}");
            return;
        }

        try
        {
            // Load the workbook; Aspose.Cells automatically detects the format
            Workbook workbook = new Workbook(inputFile);

            // Demonstrate that the workbook is loaded
            Console.WriteLine("Workbook loaded successfully.");
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);

            // (Optional) Save the workbook to verify the load succeeded
            string outputFile = "LoadedFromExcel95.xls";
            workbook.Save(outputFile);
            Console.WriteLine("Workbook saved as: " + outputFile);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
