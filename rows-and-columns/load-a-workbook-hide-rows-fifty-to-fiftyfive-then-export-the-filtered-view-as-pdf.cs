// Title: Hide rows 51‑56 in an Excel worksheet and export the visible content to PDF with Aspose.Cells for .NET
// AI Prompts: Load an XLSX file, hide rows 51 through 56 using the Cells.Rows.IsHidden property, and save the worksheet as a PDF that excludes the hidden rows (C# Aspose.Cells). | Programmatically hide a specific row range in a workbook and generate a filtered PDF output with Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells hide rows 51-56 before converting Excel to PDF in C# | C# export Excel to PDF while excluding hidden rows using Aspose.Cells | How to set row visibility and generate PDF with Aspose.Cells .NET | Hide a range of rows in an Excel workbook and save as PDF with Aspose.Cells | Aspose.Cells PDF export ignoring hidden rows in .NET
// Tags: row visibility Aspose.Cells C# | export worksheet to PDF Aspose.Cells | Cells.Rows.IsHidden property | PDF conversion without hidden rows .NET | Excel row range filtering Aspose

using System;
using System.IO;
using Aspose.Cells;

// The program loads input.xlsx, hides rows 51‑56 on the first worksheet by setting the IsHidden property, and then saves the workbook as output.pdf where the hidden rows are omitted.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return;
        }

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Hide rows 51 to 56 in Excel (zero‑based indices 50 to 55)
        for (int rowIndex = 50; rowIndex <= 55; rowIndex++)
        {
            // Use Cells.Rows collection to access row objects
            sheet.Cells.Rows[rowIndex].IsHidden = true;
        }

        try
        {
            // Export the worksheet to PDF; hidden rows are omitted from the output
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving PDF: {ex.Message}");
        }
    }
}
