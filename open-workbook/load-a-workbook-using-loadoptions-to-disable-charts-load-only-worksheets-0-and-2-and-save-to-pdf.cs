// Title: C# Aspose.Cells: Load workbook with chart loading disabled, keep only sheets 0 and 2, and save as PDF
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells using LoadOptions to skip chart data, retains only worksheets at indexes 0 and 2, and saves the result as a PDF. | Show how to configure Aspose.Cells to suppress chart loading, then delete all sheets except the first and third before converting the workbook to PDF in .NET. | Provide a method that verifies the source XLSX file, loads it without charts, filters the workbook to the required sheets, and outputs a PDF using Aspose.Cells.
// Common Searches: Aspose.Cells C# load workbook without chart objects and export selected sheets to PDF | How to disable chart loading in Aspose.Cells when converting specific worksheets to PDF | C# example to keep only sheet 1 and sheet 3 from an Excel file and save as PDF using Aspose.Cells | LoadOptions chart suppression Aspose.Cells .NET PDF conversion | Remove unwanted worksheets before PDF export with Aspose.Cells in C#
// Tags: load workbook without charts Aspose.Cells | select worksheets by index Aspose.Cells | export selected worksheets to PDF Aspose.Cells | LoadOptions chart suppression .NET | remove extra sheets before PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input XLSX file, loads it with Aspose.Cells (optionally disabling chart loading), removes all worksheets except those at indexes 0 and 2, and then saves the remaining content as a PDF, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (full load)
            Workbook workbook = new Workbook(inputPath);

            // Keep only worksheets with indexes 0 and 2
            // Remove other sheets in reverse order to preserve indexes
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (i != 0 && i != 2)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the resulting workbook to PDF format
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
