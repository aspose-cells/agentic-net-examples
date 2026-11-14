// Title: How to delete column Z, hide rows 50‑55, and save a worksheet as PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# delete column Z and hide rows 50 to 55 before exporting to PDF | Remove a specific column and hide a range of rows in an Excel file with Aspose.Cells then convert to PDF | C# code to hide rows 50‑55 and delete column Z using Aspose.Cells and save as PDF | Aspose.Cells example for deleting a column by index and hiding rows prior to PDF conversion | How to programmatically hide rows and delete a column in a workbook with Aspose.Cells .NET
// Tags: delete column by index Aspose.Cells C# | hide rows range Aspose.Cells | export worksheet to PDF Aspose.Cells | modify Excel before PDF conversion Aspose.Cells | row visibility control Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads input.xlsx, deletes column Z (index 25), hides rows 50‑55, and saves the modified workbook as output.pdf in PDF format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Delete column Z (zero‑based index 25)
            sheet.Cells.DeleteColumn(25);

            // Hide rows 50 to 55 (zero‑based indices 49 to 54)
            for (int rowIndex = 49; rowIndex <= 54; rowIndex++)
            {
                // Use the IsHidden property to hide the row
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Save the modified workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
