// Title: Auto‑fit rows that contain merged cells before saving an Excel workbook as PDF using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit row height for merged cells prior to PDF conversion | How to ensure merged cell content displays correctly in PDF with Aspose.Cells .NET | Auto‑fit rows that include merged ranges before exporting to PDF in Aspose.Cells | Adjust row heights of merged cells in an Excel workbook using Aspose.Cells C# | Aspose.Cells .NET PDF export merged cells row height issue fix
// Tags: auto-fit rows merged cells Aspose.Cells | pdf conversion row height adjustment .NET | worksheet merged range row auto-fit C# | Aspose.Cells export merged rows to PDF | row height handling merged cells Excel

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook, iterates over each merged cell area in the first worksheet, auto‑fits every row that participates in those merged ranges, and then saves the workbook as a PDF so that merged content is rendered correctly.
class Program
{
    static void Main()
    {
        try
        {
            const string inputFile = "input.xlsx";
            const string outputFile = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Iterate through all merged cell areas in the worksheet
            foreach (CellArea mergedArea in sheet.Cells.MergedCells)
            {
                // Auto‑fit each row that participates in the merged area
                for (int row = mergedArea.StartRow; row <= mergedArea.EndRow; row++)
                {
                    sheet.AutoFitRow(row);
                }
            }

            // Save the workbook as PDF; merged content will be displayed correctly
            workbook.Save(outputFile, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
