// Title: Open an Excel 95 (.xls) workbook in Aspose.Cells for .NET using LoadOptions (Excel97To2003)
// Description: Shows how to check a legacy .xls file, set LoadOptions.Version to Excel97To2003 (the nearest supported format for Excel 95), create a Workbook with those options, read the first worksheet, and optionally save the result as .xlsx.
// Keywords: Aspose.Cells | C# | LoadOptions | Excel95 | Excel97To2003 | legacy .xls | open workbook | convert to xlsx | load legacy Excel | Workbook constructor
// Common Searches: How to open Excel 95 file with Aspose.Cells .NET | LoadOptions for Excel 95 in Aspose.Cells | Convert legacy .xls to xlsx using Aspose.Cells | C# code to read old Excel workbook Aspose | Aspose.Cells support for Excel 95 format
// Developer Intent: Load a legacy Excel 95 workbook by configuring LoadOptions before instantiating the Workbook.
// Use Cases: Read data from an old Excel 95 file and display the first worksheet name. | Migrate a legacy .xls workbook to modern .xlsx for downstream processing. | Validate the existence of a legacy file and handle load errors gracefully.
// AI Prompts: Generate C# code that opens an Excel 95 (.xls) file with Aspose.Cells LoadOptions and saves it as Xlsx. | Explain why LoadFormat.Excel97To2003 is used for Excel 95 files in Aspose.Cells and how to manage format limitations. | Provide best‑practice error handling when loading legacy Excel workbooks with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Shows how to check a legacy .xls file, set LoadOptions.Version to Excel97To2003 (the nearest supported format for Excel 95), create a Workbook with those options, read the first worksheet, and optionally save the result as .xlsx.
    class LoadExcel95Example
    {
        static void Main()
        {
            try
            {
                // Path to the Excel 95 file (treated as Excel 97-2003 format, which is the closest supported)
                string filePath = "sample.xls";

                // Verify that the input file exists
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Create LoadOptions with the appropriate format (Excel97To2003 is the closest supported format)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Excel97To2003);

                // Load the workbook using the specified LoadOptions
                Workbook workbook = new Workbook(filePath, loadOptions);

                // Access the first worksheet to verify successful load
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("First worksheet name: " + sheet.Name);

                // Save the workbook to a modern format (optional)
                string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved as {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
