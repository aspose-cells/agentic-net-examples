// Title: Load an .xls workbook with Aspose.Cells LoadOptions and save it as .xlsx using C#
// Common Searches: how to load a legacy .xls file with LoadOptions in Aspose.Cells C# | convert .xls to .xlsx while controlling VBA macro import using Aspose.Cells | C# example for using Aspose.Cells LoadOptions to open Excel workbooks | Aspose.Cells workbook import with macro filtering options in .NET | save loaded .xls workbook as .xlsx with Aspose.Cells in C#
// Tags: Aspose.Cells LoadOptions workbook import | convert .xls to .xlsx Aspose.Cells C# | filter VBA projects during Excel load Aspose.Cells | handle unsigned macros Aspose.Cells LoadOptions | C# Excel file conversion with Aspose.Cells | save workbook to different format .NET Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to verify an input .xls file, create a LoadOptions object, load the workbook with Aspose.Cells, ensure the output directory exists, and then save the workbook as an .xlsx file, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "InputWorkbook.xls";
            string outputPath = "OutputWorkbook.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with default options (VBA macro filtering not supported in this version)
            LoadOptions loadOptions = new LoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in the desired format
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
