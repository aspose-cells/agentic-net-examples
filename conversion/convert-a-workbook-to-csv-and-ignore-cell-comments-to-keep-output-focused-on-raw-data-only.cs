// Title: Convert an Excel workbook to CSV while excluding cell comments with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to CSV without including comments | How to save .xlsx as .csv using Aspose.Cells while ignoring cell notes | C# convert Excel file to CSV format and omit comments with TxtSaveOptions | Aspose.Cells CSV conversion ignore cell comments in .NET application | Save Excel workbook as CSV in C# and prevent comment export using Aspose.Cells
// Tags: Aspose.Cells CSV export ignoring comments | TxtSaveOptions for CSV output in .NET | C# Excel to CSV conversion using Aspose.Cells | exclude cell comments during CSV save | SaveFormat.Csv with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example checks for the source .xlsx file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions with SaveFormat.Csv (comments are omitted by default), saves the workbook as a CSV file, and handles any errors, providing console feedback.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure TXT save options for CSV output
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv);
            // Note: ExportComments property is not available in the current API version.
            // By default, comments are not exported in CSV format.

            // Save the workbook as a CSV file
            workbook.Save(outputPath, txtOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
