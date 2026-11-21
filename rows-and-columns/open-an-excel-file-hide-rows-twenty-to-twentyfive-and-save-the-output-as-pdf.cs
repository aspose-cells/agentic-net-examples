// Title: Hide rows 20‑25 in an Excel workbook and export to PDF using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide rows 20 to 25 before saving as PDF | How to programmatically hide a range of rows in an Excel worksheet using Aspose.Cells | Convert Excel workbook to PDF while keeping specific rows hidden with Aspose.Cells .NET | C# code to set IsHidden property for rows 20‑25 in Aspose.Cells workbook | Export hidden rows Excel to PDF using Aspose.Cells SaveFormat.Pdf
// Tags: hide specific rows Aspose.Cells C# | set row IsHidden property Aspose.Cells | export workbook to PDF Aspose.Cells | Excel to PDF conversion with hidden rows | row visibility control Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file with Aspose.Cells, hides rows 20 through 25 on the first worksheet by setting each row's IsHidden property, and then saves the modified workbook as a PDF using SaveFormat.Pdf, with basic file‑existence checking and exception handling.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the Excel file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide rows 20 to 25 (zero‑based indices 19 to 24)
            for (int rowIndex = 19; rowIndex <= 24; rowIndex++)
            {
                // Use Cells.Rows to access the Row object and set its IsHidden property
                sheet.Cells.Rows[rowIndex].IsHidden = true;
            }

            // Save the modified workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
