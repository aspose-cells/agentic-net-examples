// Title: Export a specific Excel worksheet to CSV while trimming leading empty rows and columns using Aspose.Cells for .NET (C#)
// Common Searches: how to export only one sheet from an .xlsx to CSV with Aspose.Cells C# and ignore leading blank rows | Aspose.Cells C# save active worksheet as CSV without empty top rows and left columns | C# convert a selected Excel worksheet to CSV trimming leading empty rows and columns using Aspose.Cells | export chosen sheet to CSV in .NET while removing initial blank rows and columns | Aspose.Cells CSV export options to skip leading empty rows in C#
// Tags: export specific worksheet to CSV Aspose.Cells | trim leading blank rows during CSV export C# | save active sheet as CSV .NET | Aspose.Cells CSV export options | remove empty rows and columns when converting Excel to CSV | C# worksheet selection for CSV output

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, selects a worksheet by name or index, makes it the active sheet, and then saves only that active worksheet as a CSV file. During the CSV save Aspose.Cells automatically omits leading empty rows and columns, producing a compact CSV output. Exceptions are caught and reported.
class ExportWorksheetToCsv
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);

            // Select the worksheet to export (by name or index).
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0];

            // Make the target worksheet the active one.
            workbook.Worksheets.ActiveSheetIndex = worksheet.Index;

            // Save the active worksheet as CSV. When saving as CSV, only the active sheet is exported.
            workbook.Save(outputPath, SaveFormat.Csv);

            Console.WriteLine($"Worksheet '{worksheet.Name}' exported successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
