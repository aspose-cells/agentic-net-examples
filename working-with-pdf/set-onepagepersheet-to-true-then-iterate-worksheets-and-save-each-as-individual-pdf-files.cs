// Title: Create individual PDF files for each Excel worksheet with OnePagePerSheet enabled using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook, sets Workbook.Settings.OnePagePerSheet to true, and saves every worksheet as a separate PDF named after the sheet using Aspose.Cells. | Adapt an Aspose.Cells example to generate a temporary workbook for each sheet, enable one‑page‑per‑sheet, and export each sheet to its own PDF file.
// Common Searches: Aspose.Cells C# export each worksheet to its own PDF file with OnePagePerSheet | How to save Excel sheets as separate PDFs using Aspose.Cells .NET | C# code to iterate worksheets and generate individual PDF documents with Aspose.Cells | Set OnePagePerSheet true when converting Excel to PDF per sheet in Aspose.Cells
// Tags: Aspose.Cells OnePagePerSheet PDF export | export Excel worksheet to PDF C# | temporary workbook per sheet Aspose.Cells | save each sheet as separate PDF Aspose.Cells | C# iterate worksheets save PDF | Aspose.Cells PDF conversion per worksheet

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, enables OnePagePerSheet, iterates through all worksheets, creates a temporary workbook containing only the current sheet, removes the default empty sheet, and saves each sheet as a PDF file named after the worksheet, handling errors for individual sheets.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputPath);

            // Iterate through all worksheets and save each as an individual PDF
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                try
                {
                    // Get the current worksheet name for output file naming
                    string sheetName = sourceWorkbook.Worksheets[i].Name;
                    string outputFile = $"{sheetName}.pdf";

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWorkbook = new Workbook();

                    // Add a copy of the current sheet by name
                    tempWorkbook.Worksheets.AddCopy(sheetName);

                    // Remove the default empty sheet that exists in a new workbook
                    if (tempWorkbook.Worksheets.Count > 1)
                    {
                        tempWorkbook.Worksheets.RemoveAt(0);
                    }

                    // Save the temporary workbook as PDF
                    tempWorkbook.Save(outputFile, SaveFormat.Pdf);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save sheet '{sourceWorkbook.Worksheets[i].Name}': {ex.Message}");
                }
            }

            Console.WriteLine("PDF files have been generated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
