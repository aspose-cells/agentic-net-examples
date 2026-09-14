// Title: Set worksheet print area to its used range and export the sheet as PDF using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, assign the MaxDisplayRange of the first worksheet to PageSetup.PrintArea, and save the workbook as a PDF with Aspose.Cells. | Programmatically define the print area based on the worksheet's used range and generate a PDF file from the workbook in C#.
// Common Searches: Aspose.Cells C# set print area to used range before PDF conversion | How to export an Excel worksheet to PDF with a custom print area using Aspose.Cells | C# code to set PageSetup.PrintArea to MaxDisplayRange in Aspose.Cells | Define print area programmatically in Aspose.Cells .NET and save as PDF | Export first sheet of workbook to PDF after setting print area with Aspose.Cells
// Tags: set worksheet print area MaxDisplayRange | export worksheet to PDF Aspose.Cells | page setup print area C# | save workbook as PDF custom print area | Aspose.Cells used range print area

using System;
using System.IO;
using Aspose.Cells;

// Loads an Excel file, sets the first worksheet's print area to its used range via MaxDisplayRange, and saves the workbook as a PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Define input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get the used range as a string (e.g., "A1:D10")
            string printArea = sheet.Cells.MaxDisplayRange.RefersTo;

            // Set the print area for the worksheet
            sheet.PageSetup.PrintArea = printArea;

            // Export the worksheet to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
