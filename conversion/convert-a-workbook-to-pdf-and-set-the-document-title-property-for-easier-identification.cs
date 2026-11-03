// Title: Convert an .xlsx workbook to PDF and set the document title property using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert Excel to PDF and set built‑in document title | How to add PDF title metadata when exporting .xlsx with Aspose.Cells | Aspose.Cells example for handling password‑protected Excel during PDF conversion | Save Excel workbook as PDF with custom title property using Aspose.Cells C# | Load password protected workbook in Aspose.Cells and export to PDF with metadata
// Tags: Aspose.Cells Excel to PDF conversion with title metadata | set built‑in document title property Aspose.Cells | password protected workbook PDF export Aspose.Cells | SaveFormat.Pdf usage Aspose.Cells C# | LoadOptions password handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an .xlsx file with Aspose.Cells, assigns a custom Title to the workbook's built‑in document properties, and saves the workbook as a PDF. It includes logic to detect password‑protected files, reload them using LoadOptions with a password, and then perform the same PDF conversion while handling errors gracefully.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Verify that the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (no password). If the file is password‑protected,
                // a CellsException will be thrown and handled in the catch block below.
                Workbook workbook = new Workbook(inputPath);

                // Set a title property for easier identification
                workbook.BuiltInDocumentProperties.Title = "My Workbook Title";

                // Save the workbook as a PDF file
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files
                Console.WriteLine("The workbook appears to be password protected.");

                try
                {
                    // Replace "your_password" with the actual password
                    LoadOptions loadOptions = new LoadOptions { Password = "your_password" };
                    Workbook workbook = new Workbook(inputPath, loadOptions);
                    workbook.Save(outputPath, SaveFormat.Pdf);
                    Console.WriteLine($"Workbook successfully saved as PDF using password: {outputPath}");
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to load workbook with password: {innerEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
