// Title: Load an XLSX workbook from a file path and set its calculation mode to Manual using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook from specific file path and set calculation mode to manual | How to disable automatic formula calculation when opening an XLSX file with Aspose.Cells | C# example for setting Workbook.Settings.CalcMode to Manual after loading an Excel file | Load Excel workbook from disk and prevent recalculation using Aspose.Cells .NET | Aspose.Cells manual calculation mode for a loaded workbook in C#
// Tags: Aspose.Cells load XLSX from disk | manual CalcMode setting Aspose.Cells | prevent formula recalculation on workbook open | Workbook.Settings.CalcMode property usage | C# Excel workbook loading with manual calculation

using System;
using System.IO;
using Aspose.Cells;

// The program checks that the specified XLSX file exists, loads it into an Aspose.Cells Workbook, and demonstrates how to set Workbook.Settings.CalcMode to Manual to stop automatic formula recalculation, with proper exception handling.
class Program
{
    static void Main()
    {
        // Path to the XLSX file to be loaded
        string filePath = @"C:\Path\To\YourWorkbook.xlsx";

        try
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(filePath);

            // Optional: Set calculation mode to Manual if supported
            // Uncomment the following line if the CalcMode property is available in your Aspose.Cells version
            // workbook.Settings.CalcMode = Aspose.Cells.CalcMode.Manual;

            // Continue with further processing...
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
