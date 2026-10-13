// Title: Hide columns D through G in an Excel worksheet and export the sheet to PDF with Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file, set the IsHidden property for columns D to G on the first worksheet, and save the result as a PDF using Aspose.Cells in C#. | Write a C# console app that checks for an input Excel file, hides a specific column range, and generates a PDF output with Aspose.Cells. | Create code that programmatically hides multiple columns in a workbook and performs a PDF conversion with Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# hide column range D-G before saving as PDF | how to programmatically hide multiple Excel columns and export to PDF using Aspose.Cells | C# example to hide columns D through G in a worksheet and convert to PDF with Aspose.Cells
// Tags: hide columns Aspose.Cells .NET | export worksheet to PDF Aspose.Cells | column visibility Excel Aspose.Cells | save workbook as PDF after hiding columns | programmatic column hide C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads input.xlsx with Aspose.Cells, hides columns D through G on the first worksheet, ensures the output folder exists, and saves the workbook as output.pdf.
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
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Hide columns D (index 3) through G (index 6) individually
            for (int col = 3; col <= 6; col++)
            {
                // Set the column's IsHidden property to true
                sheet.Cells.Columns[col].IsHidden = true;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
