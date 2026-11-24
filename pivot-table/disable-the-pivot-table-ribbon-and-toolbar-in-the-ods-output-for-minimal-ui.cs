// Title: Export an Excel workbook to OpenDocument Spreadsheet (ODS) using Aspose.Cells for .NET
// Common Searches: C# how to save an Excel file as ODS with Aspose.Cells | Aspose.Cells OdsSaveOptions usage example in .NET | convert .xlsx to .ods programmatically using C# | export workbook to OpenDocument Spreadsheet format with Aspose.Cells | Aspose.Cells .NET save workbook as ODS without data loss
// Tags: ODS export options in Aspose.Cells | C# convert XLSX to OpenDocument Spreadsheet | Aspose.Cells save workbook as ODS format | programmatic ODS generation using Aspose.Cells .NET | OpenDocument Spreadsheet export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

// The example loads an existing Excel workbook (or creates a new one if missing), configures OdsSaveOptions for the ODS format, and saves the workbook as an OpenDocument Spreadsheet file. No settings are applied to hide the pivot table ribbon or toolbar in the resulting ODS document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Ensure the input file exists; otherwise create a new workbook.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found. Creating a new workbook.");
                workbook = new Workbook();
                // Optionally add sample data here.
            }

            // Configure ODS save options.
            OdsSaveOptions odsOptions = new OdsSaveOptions(SaveFormat.Ods);

            // Save the workbook as ODS with the specified options.
            workbook.Save(outputPath, odsOptions);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
