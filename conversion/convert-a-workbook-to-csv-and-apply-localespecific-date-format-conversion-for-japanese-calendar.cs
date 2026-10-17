// Title: Convert an Excel workbook to CSV while applying Japanese calendar date formatting using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to CSV with Japanese calendar dates in C# | How to set CultureInfo to ja-JP before saving workbook as CSV using Aspose.Cells | C# convert .xlsx to .csv applying Japanese locale date format with Aspose.Cells | Save workbook as CSV with Japanese date formatting using Aspose.Cells .NET API | Aspose.Cells CSV output Japanese date format example C#
// Tags: Aspose.Cells CSV export Japanese locale | set workbook CultureInfo ja-JP Aspose.Cells | Excel to CSV conversion locale-specific formatting | C# Aspose.Cells Japanese calendar date format | save workbook as CSV using CultureInfo

using Aspose.Cells;
using System;
using System.Globalization;
using System.IO;

// The example loads an existing XLSX file, assigns the Japanese culture (ja-JP) to the workbook's Settings.CultureInfo to enforce Japanese calendar date formatting, and then saves the workbook as a CSV file, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            var workbook = new Workbook(inputPath);

            // Apply Japanese locale for culture‑specific formatting
            workbook.Settings.CultureInfo = new CultureInfo("ja-JP");

            // Save the workbook as CSV
            workbook.Save(outputPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
