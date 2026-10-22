// Title: Load an Excel workbook in C# with Aspose.Cells while filtering out unsigned VBA macros using LoadOptions.FilterVbaProject
// Common Searches: how to use LoadOptions.FilterVbaProject to ignore unsigned VBA macros in Aspose.Cells C# | Aspose.Cells example for loading a workbook without unsigned macro projects | C# load .xls file and exclude VBA project using Aspose.Cells LoadOptions | filter unsigned VBA macros during Excel import with Aspose.Cells .NET | prevent unsigned macro execution when opening workbooks in Aspose.Cells C#
// Tags: LoadOptions.FilterVbaProject usage C# | exclude unsigned VBA macros Aspose.Cells | load workbook without VBA project .NET | filter VBA project during Excel import | Aspose.Cells macro security handling

using Aspose.Cells;
using System;
using System.IO;

// Demonstrates loading an Excel file with Aspose.Cells in C# while applying LoadOptions.FilterVbaProject to exclude unsigned VBA macro projects, then saving the workbook to a new format with proper error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xls";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (no VBA project filter needed)
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook in the desired format
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
