// Title: Convert an Excel .xlsx workbook to CSV while keeping empty cells as empty strings using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export .xlsx to .csv preserving blank cells | How to keep empty cells when converting Excel workbook to CSV with Aspose.Cells | C# save Excel file as CSV without dropping empty columns using Aspose.Cells | Maintain column count during .xlsx to .csv conversion in .NET Aspose.Cells | Aspose.Cells SaveFormat.Csv keep empty strings in output file
// Tags: Aspose.Cells CSV export blank cell handling | C# workbook to CSV conversion SaveFormat.Csv | retain empty cells during Excel to CSV conversion | ensure column consistency when converting .xlsx to .csv | Aspose.Cells .NET CSV output empty string handling

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies that the input.xlsx file exists, loads it into an Aspose.Cells Workbook, and saves it as output.csv using SaveFormat.Csv. Aspose.Cells automatically writes empty cells as empty strings, preserving column counts, and the code includes basic exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as CSV. Aspose.Cells automatically preserves empty cells.
            workbook.Save(outputPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
