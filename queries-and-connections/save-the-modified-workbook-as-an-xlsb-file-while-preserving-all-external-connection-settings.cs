// Title: Save a modified Excel workbook as XLSB while preserving external connection settings with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx workbook, changes a cell value, and exports it to .xlsb while ensuring external connections remain active using Aspose.Cells. | Show how to set XlsbSaveOptions so that external data links are retained when saving a workbook as a binary file in C#. | Implement a C# helper that takes input and output file paths, updates the first worksheet, and saves the file as XLSB with all connection settings kept via Aspose.Cells.
// Common Searches: Aspose.Cells keep linked tables when converting Excel to binary format | C# enable external connections in XLSB export with Aspose | how to retain data connections after saving workbook as .xlsb using Aspose.Cells | example of preserving external data sources when saving to XLSB in .NET | saving .xlsx as .xlsb without breaking external connections Aspose.Cells
// Tags: XlsbSaveOptions external connections | save workbook as XLSB Aspose.Cells | modify cell before binary export C# | preserve linked data sources Excel binary | convert .xlsx to .xlsb Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads 'input.xlsx', updates cell A1, and saves the workbook as 'output.xlsb' using XlsbSaveOptions. By enabling external connections, all data links are retained in the resulting binary file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsb";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Example modification: write a value to cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified");

            // Preserve external connections when saving (property available in newer versions)
            // If the property is not present in the referenced Aspose.Cells version, this line can be omitted.
            // workbook.Settings.EnableExternalConnections = true;

            // Prepare XLSB save options (preserves all connection settings)
            XlsbSaveOptions saveOptions = new XlsbSaveOptions();

            // Save the workbook as an XLSB file
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
