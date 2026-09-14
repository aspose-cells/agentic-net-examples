// Title: Export each worksheet's paper size and dimensions to a CSV file with Aspose.Cells for .NET
// AI Prompts: Write a C# program that iterates over every worksheet in a workbook, reads the PageSetup paper size, width, and height, and writes these values together with the sheet name into a CSV file using Aspose.Cells. | Modify the export utility to accept the input Excel file path and the output CSV path as command‑line arguments, providing default values when they are omitted. | Enhance the CSV output to also record the page orientation (portrait or landscape) for each worksheet alongside the existing paper size and dimensions.
// Common Searches: how to list paper size and dimensions of all sheets in an Excel workbook using Aspose.Cells C# | C# export Excel worksheet page setup properties to CSV with Aspose.Cells | retrieve worksheet paper width and height in inches via Aspose.Cells API | command line tool for extracting Excel sheet page setup to CSV in .NET
// Tags: Aspose.Cells page setup CSV generation | C# get worksheet paper dimensions inches | batch export Excel sheet page properties | extract page orientation Aspose.Cells | command line workbook page setup exporter

using System;
using System.IO;
using Aspose.Cells;

// // Exports each worksheet's name, paper size enum, and paper width/height in inches to a CSV file using Aspose.Cells.
class ExportWorksheetPaperDimensions
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Prepare the CSV output file (replace with desired output path)
        string csvPath = "WorksheetPaperDimensions.csv";
        using (StreamWriter writer = new StreamWriter(csvPath))
        {
            // Write CSV header
            writer.WriteLine("WorksheetName,PaperSize,PaperWidthInches,PaperHeightInches");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the page setup of the worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Get paper size enum name
                string paperSizeName = pageSetup.PaperSize.ToString();

                // Get paper width and height in inches (if not set, defaults to 0)
                double paperWidth = pageSetup.PaperWidth;
                double paperHeight = pageSetup.PaperHeight;

                // Write a line for the current worksheet
                writer.WriteLine($"{EscapeCsv(sheet.Name)},{paperSizeName},{paperWidth},{paperHeight}");
            }
        }

        Console.WriteLine($"Worksheet paper dimensions have been exported to '{csvPath}'.");
    }

    // Helper method to escape CSV fields that may contain commas or quotes
    private static string EscapeCsv(string field)
    {
        if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
        {
            field = field.Replace("\"", "\"\"");
            return $"\"{field}\"";
        }
        return field;
    }
}
