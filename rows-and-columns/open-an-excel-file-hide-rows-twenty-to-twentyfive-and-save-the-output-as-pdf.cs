// Title: Hide rows 20‑25 in an Excel worksheet and export the workbook as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing .xlsx file, set rows 20‑25 to hidden, and save the workbook as a PDF with Aspose.Cells in C#. | Programmatically conceal a range of rows in a worksheet and generate a PDF output using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells hide rows 20 to 25 before converting to PDF | How to hide specific rows in an Excel file and export to PDF using Aspose.Cells .NET | Aspose.Cells example hide row range and save as PDF | Convert Excel to PDF while preserving hidden rows with Aspose.Cells C#
// Tags: Aspose.Cells hide rows C# | Aspose.Cells export to PDF C# | Excel row visibility Aspose.Cells | C# convert hidden rows Excel to PDF | Aspose.Cells IsHidden property

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, hides rows 20‑25 on the first worksheet using the IsHidden property, and saves the modified workbook as output.pdf.
class HideRowsAndExportPdf
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify by name/index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 20 to 25 (Aspose.Cells uses zero‑based indexing)
            for (int rowIndex = 19; rowIndex <= 24; rowIndex++)
            {
                // Use IsHidden property to hide the row
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
