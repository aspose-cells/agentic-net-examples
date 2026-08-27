// Title: How to delete column Q from an Excel worksheet and save the result as a PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Delete column Q (zero‑based index 16) from a workbook, update any dependent formulas, and export the first worksheet to PDF with Aspose.Cells in C#. | Remove a specific column from an Excel file and convert the modified sheet to PDF using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# delete column by index and keep formula references before PDF export | C# code to remove column Q from an Excel file and generate a PDF with Aspose.Cells | How to delete a column in an Excel workbook using Aspose.Cells and then save as PDF | Update formulas after deleting a column in Aspose.Cells before converting to PDF | Convert modified worksheet to PDF after column removal using Aspose.Cells .NET
// Tags: delete column Aspose.Cells C# | formula update after column deletion Aspose.Cells | save worksheet as PDF Aspose.Cells .NET | remove column Q by index Aspose.Cells | convert modified Excel workbook to PDF using Aspose

using System;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, deletes column Q (index 16) while updating formulas, and then saves the resulting worksheet directly as a PDF file.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputPath = "input.xlsx";

        // Load the workbook from the file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(inputPath);

        // Get the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete column Q.
        // Column indices are zero‑based, so column Q is index 16 (A=0, B=1, ..., Q=16).
        // The second parameter 'true' updates any formulas that reference the deleted column.
        worksheet.Cells.DeleteColumn(16, true);

        // Save the modified workbook as a PDF file (uses Workbook.Save(string, SaveFormat))
        string outputPath = "output.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
    }
}
