// Title: Automatically detect Excel workbook format by using only the file path with Aspose.Cells Workbook constructor in C#
// Common Searches: Aspose.Cells C# load workbook without specifying format using file path | How to let Aspose.Cells infer Excel file type from its extension when opening | Create a new workbook automatically when source file does not exist Aspose.Cells | Save Aspose.Cells workbook and have format determined by output file extension | C# Aspose.Cells auto-detect input workbook format from .xls or .xlsx path
// Tags: auto format detection via Workbook constructor Aspose.Cells | load Excel file by path without format argument | save workbook with inferred extension Aspose.Cells | fallback create new workbook when file missing Aspose.Cells | file extension based format inference Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example checks whether the input Excel file exists; if it does, a Workbook is instantiated with just the file path, letting Aspose.Cells automatically determine the source format (XLS, XLSX, etc.). If the file is missing, a new default workbook is created. The workbook is then saved to a target path, where the output format is inferred from the file extension.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input and output workbooks
            string inputPath = @"C:\Data\SourceWorkbook.xls";
            string outputPath = @"C:\Data\ConvertedWorkbook.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if it exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath); // format auto‑detected
                }
                else
                {
                    Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook.");
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Save the workbook; format inferred from the file extension
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
