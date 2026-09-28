// Title: Load an Excel workbook, keep only worksheets 1 and 3, and export to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, removes every worksheet except the first and third, and saves the result as a PDF. | Write a C# program using Aspose.Cells to load a workbook, filter out unwanted sheets, and export the remaining sheets (1 and 3) to a PDF document. | Create a C# example that demonstrates how to retain specific worksheets in an Excel file and convert them to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# export only selected worksheets to PDF | How to delete specific sheets in an Excel workbook before PDF conversion using Aspose.Cells | C# keep first and third worksheet when saving Excel as PDF with Aspose.Cells | Remove all sheets except sheet 1 and 3 in Aspose.Cells before PDF export | Aspose.Cells load workbook and save only certain sheets as PDF in .NET
// Tags: Aspose.Cells selective worksheet PDF conversion | C# remove Excel sheets before PDF export | Aspose.Cells keep specific worksheets | Excel to PDF conversion with sheet filtering C# | Aspose.Cells workbook sheet removal

using Aspose.Cells;
using System;
using System.IO;

// The program verifies the input Excel file, loads it with Aspose.Cells, removes all worksheets except the first and third (zero‑based indexes 0 and 2), and then saves the remaining content as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (standard loading; LightCells not required for this example)
            Workbook workbook = new Workbook(inputPath);

            // Keep only worksheets 1 and 3 (zero‑based indexes 0 and 2)
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (i != 0 && i != 2)
                    workbook.Worksheets.RemoveAt(i);
            }

            // Save the resulting workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
