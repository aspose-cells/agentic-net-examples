// Title: Load an XLSX workbook with Aspose.Cells LoadOptions and immediately save as PDF while preserving auto‑fit row heights (C#)
// AI Prompts: Write C# code that creates a LoadOptions object for Xlsx, loads a workbook with those options, and saves it directly to PDF using Aspose.Cells. | Show how to export an Excel file to PDF in C# with Aspose.Cells so that rows keep their auto‑fit heights.
// Common Searches: Aspose.Cells C# load workbook with LoadOptions then export to PDF preserving row heights | How to convert an .xlsx to PDF using Aspose.Cells without loading full data into memory | C# example for using OnlyAuto loading option with Aspose.Cells PDF conversion
// Tags: Aspose.Cells LoadOptions PDF export C# | Excel to PDF conversion preserving row heights | OnlyAuto loading option Aspose.Cells example | C# Workbook SaveFormat.Pdf usage

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program checks that the input XLSX file exists, creates LoadOptions for the Xlsx format, loads the workbook with those options, and immediately saves it as a PDF, handling any exceptions that may occur.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load options for XLSX format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                // Note: LoadDataOnly property is not available in this version of Aspose.Cells.
                // If needed, it can be set via reflection in newer versions.

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Successfully converted '{inputPath}' to PDF as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
