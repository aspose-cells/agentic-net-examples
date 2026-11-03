// Title: C# – Remove empty rows and columns from an Excel workbook and save it as ODS using Aspose.Cells
// Common Searches: how to delete empty rows and columns from an Excel file with Aspose.Cells C# | Aspose.Cells C# remove blank rows before exporting to ODS | delete blank columns in each worksheet using Aspose.Cells .NET | convert cleaned XLSX to ODS with Aspose.Cells SaveFormat Ods | C# code to clean Excel workbook and save as ODS format
// Tags: Aspose.Cells DeleteBlankRows usage | Aspose.Cells DeleteBlankColumns usage | save workbook as ODS Aspose.Cells | clean Excel worksheet before ODS export | non‑obsolete SaveFormat Ods | remove empty rows columns .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing XLSX file with Aspose.Cells, iterates through all worksheets, removes rows and columns that contain no data using DeleteBlankRows and DeleteBlankColumns, and then saves the cleaned workbook as an ODS file using the non‑obsolete SaveFormat.Ods enum, with basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.ods";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and remove empty rows and columns
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Remove rows that contain no data
                sheet.Cells.DeleteBlankRows();

                // Remove columns that contain no data
                sheet.Cells.DeleteBlankColumns();
            }

            // Save the cleaned workbook in ODS format (use the non‑obsolete enum value)
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
