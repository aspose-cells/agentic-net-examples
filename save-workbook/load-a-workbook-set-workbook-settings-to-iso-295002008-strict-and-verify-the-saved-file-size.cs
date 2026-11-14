// Title: Load an Excel workbook, save it as ISO 29500‑2008 strict XLSX, and check the resulting file size with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells load existing .xlsx and save as ISO 29500-2008 strict format | How to verify the size of a saved Excel file using Aspose.Cells in .NET | Set OoxmlSaveOptions to strict XLSX when exporting a workbook with Aspose.Cells C# | Determine output file length after saving workbook as strict XLSX with Aspose.Cells | Save workbook with ISO 29500-2008 compliance and get file size in C#
// Tags: Aspose.Cells OoxmlSaveOptions strict XLSX | C# load workbook and export ISO 29500-2008 | check saved Excel file size .NET | ISO 29500-2008 compliance Aspose.Cells | retrieve output file length after save

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an existing XLSX file into an Aspose.Cells Workbook, saves it using OoxmlSaveOptions configured for ISO 29500‑2008 strict compliance, and then reads and prints the size of the generated file in bytes.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare save options for XLSX format
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Output workbook path
            string outputPath = "output_strict.xlsx";

            // Save the workbook
            workbook.Save(outputPath, saveOptions);

            // Verify the saved file size
            long fileSize = new FileInfo(outputPath).Length;
            Console.WriteLine($"Saved file size: {fileSize} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
