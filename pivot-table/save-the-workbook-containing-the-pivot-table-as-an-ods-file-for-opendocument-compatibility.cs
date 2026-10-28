// Title: Convert an Excel workbook with a pivot table to OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells save workbook containing pivot table as .ods file | How to export Excel pivot table to OpenDocument format with Aspose.Cells .NET | Aspose.Cells convert password‑protected XLSX with pivot tables to ODS in C# | Saving pivot table workbook to ODS using SaveFormat.Ods in Aspose.Cells | Load Excel file and export to ODS while preserving pivot tables in .NET
// Tags: Aspose.Cells save workbook as ODS | export pivot table to OpenDocument Spreadsheet | convert XLSX to ODS with Aspose.Cells | handle password protected workbook Aspose.Cells | SaveFormat.Ods usage in C# | preserve pivot tables when converting to ODS

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, verifies its presence, and uses Aspose.Cells to save the workbook—including any pivot tables—to an OpenDocument Spreadsheet (ODS) file. It demonstrates handling of password‑protected workbooks and general exception management.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.ods";

        try
        {
            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook. If the file is password‑protected, a CellsException will be thrown.
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook in OpenDocument Spreadsheet (ODS) format.
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook successfully saved as {outputPath}");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected files or other Cells-specific errors.
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is password protected. Provide a password to open it.");
                // Example of loading with a password (uncomment and set the password if known):
                // string password = "yourPassword";
                // LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                // Workbook workbook = new Workbook(inputPath, loadOptions);
                // workbook.Save(outputPath, SaveFormat.Ods);
            }
            else
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // General exception handling.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
