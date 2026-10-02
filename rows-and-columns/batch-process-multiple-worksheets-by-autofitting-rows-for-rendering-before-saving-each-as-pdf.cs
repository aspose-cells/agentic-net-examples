// Title: Auto‑fit rows in each worksheet and export them as individual PDFs using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through all worksheets in an Excel workbook, calls AutoFitRows on each sheet, and saves each sheet as a separate PDF with Aspose.Cells. | Update the sample program to set the PDF page orientation to Landscape while still adjusting row heights before exporting each worksheet to PDF. | Write a reusable C# method that takes an Excel file path, automatically adjusts row heights on every worksheet, and returns a list of the created PDF file paths.
// Common Searches: how to adjust row heights automatically and generate a PDF for each Excel worksheet using Aspose.Cells in C# | batch conversion of Excel sheets to PDF after resizing rows with Aspose.Cells .NET | C# method to copy a worksheet to a new workbook, delete the placeholder sheet, and create a PDF file | export individual Excel worksheets to separate PDF files without extra blank pages using Aspose.Cells
// Tags: row height auto adjustment Aspose.Cells | save sheet as PDF C# | temporary workbook per sheet Aspose.Cells | remove placeholder sheet after copy Aspose.Cells | multiple worksheet PDF export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the input.xlsx file, loads it into a Workbook, loops through each worksheet, automatically adjusts row heights with AutoFitRows, copies the sheet into a temporary workbook, removes the default empty sheet, and saves the temporary workbook as a PDF named with the original file base and sheet name, handling any exceptions that occur.
class Program
{
    static void Main()
    {
        const string inputFile = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: The file \"{inputFile}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(inputFile);

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet sheet in sourceWorkbook.Worksheets)
            {
                // Auto‑fit all rows in the current worksheet for proper rendering
                sheet.AutoFitRows();

                // Create a temporary workbook that will contain only the current sheet
                Workbook tempWorkbook = new Workbook();

                // Copy the current sheet into the temporary workbook by name
                tempWorkbook.Worksheets.AddCopy(sheet.Name);

                // Remove the default empty sheet that Aspose.Cells creates
                if (tempWorkbook.Worksheets.Count > 1)
                {
                    tempWorkbook.Worksheets.RemoveAt(0);
                }

                // Build the output PDF file name (e.g., input_Sheet1.pdf)
                string pdfFileName = $"{Path.GetFileNameWithoutExtension(inputFile)}_{sheet.Name}.pdf";

                // Save the temporary workbook (which now contains only the current sheet) as PDF
                tempWorkbook.Save(pdfFileName, SaveFormat.Pdf);
                Console.WriteLine($"Saved PDF: {pdfFileName}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
