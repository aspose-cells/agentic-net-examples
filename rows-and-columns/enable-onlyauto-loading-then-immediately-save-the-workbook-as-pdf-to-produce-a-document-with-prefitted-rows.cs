// Title: Enable LoadDataOnly and auto‑fit rows before exporting an Excel workbook to PDF using Aspose.Cells for .NET
// Common Searches: how to use LoadOptions.LoadDataOnly with Aspose.Cells to speed up Excel loading in C# | auto fit all rows in a workbook before saving as PDF with Aspose.Cells .NET | convert an XLSX file to PDF while preserving row heights using Aspose.Cells | Aspose.Cells C# example for loading workbook with data‑only mode and exporting to PDF
// Tags: LoadDataOnly option Aspose.Cells | auto fit rows before PDF export | Excel to PDF conversion C# Aspose.Cells | fast workbook loading Aspose.Cells | worksheet AutoFitRows method | save workbook as PDF Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input XLSX file, creates LoadOptions (optionally enabling LoadDataOnly), loads the workbook, auto‑fits rows in every worksheet, and saves the result directly as a PDF, handling any exceptions that may occur.
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

            // Load options – if the used Aspose.Cells version supports LoadDataOnly, enable it.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // Uncomment the following line if the LoadDataOnly property is available in your version.
            // loadOptions.LoadDataOnly = true;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Auto‑fit rows in each worksheet so they are pre‑fitted before saving
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitRows();
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Successfully saved PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
