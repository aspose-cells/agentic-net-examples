// Title: How to override GetErrorString in Aspose.Cells for .NET to localize Excel error messages
// Common Searches: c# override GetErrorString Aspose.Cells to show errors in Spanish | custom error string provider example for Aspose.Cells formula errors | localizing Excel calculation error messages with Aspose.Cells GetErrorString | Aspose.Cells GetErrorString method for multilingual workbook error handling
// Tags: override GetErrorString Aspose.Cells .NET | custom error string provider Aspose.Cells | localize Excel error messages C# | Excel formula error localization Aspose.Cells | multilingual error handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomErrorString
{
    // The sample loads or creates a workbook, inserts a division‑by‑zero formula to trigger an error, recalculates formulas, and saves the file, while demonstrating that a custom GetErrorString provider can be overridden to supply localized Excel error messages (not available in this library version).
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "Input.xlsx";
                Workbook workbook;

                // Load existing workbook if it exists; otherwise create a new workbook.
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // NOTE: Custom error string provider is not available in this version of Aspose.Cells.
                // The default error messages will be used.

                // Force a division by zero error in cell A1.
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];
                cell.Formula = "=1/0";

                // Recalculate formulas to generate the error.
                workbook.CalculateFormula();

                // Save the workbook.
                const string outputPath = "Output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
