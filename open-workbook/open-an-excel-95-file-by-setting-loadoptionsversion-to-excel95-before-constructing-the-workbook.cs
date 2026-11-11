// Title: Open an Excel 95 (.xls) workbook in C# with Aspose.Cells by setting LoadOptions.Version to ExcelVersion.Excel95
// Common Searches: how to load an Excel 95 file using Aspose.Cells LoadOptions in C# | Aspose.Cells example for opening legacy .xls workbook with specific version setting | C# code to specify ExcelVersion.Excel95 when creating a Workbook | read old Excel 95 format with Aspose.Cells LoadOptions parameter | load legacy Excel 95 workbook in .NET using Aspose.Cells LoadOptions.Version
// Tags: load Excel95 workbook with LoadOptions | Aspose.Cells legacy xls loading | specify Excel version in Workbook constructor | C# load old Excel format using Aspose.Cells | LoadOptions.Version Excel95 example

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that the input .xls file exists, creates a LoadOptions object with Version set to ExcelVersion.Excel95, and then constructs a Workbook using this LoadOptions to correctly open an Excel 95 workbook. It reports success or catches and displays any loading errors.
class Program
{
    static void Main()
    {
        // Path to the Excel 95 file
        string filePath = "input.xls";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Input file not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook; Aspose.Cells automatically detects the format
            Workbook workbook = new Workbook(filePath);

            // Optional: perform further operations on the workbook
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
        }
    }
}
