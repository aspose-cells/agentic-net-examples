// Title: Export an Aspose.Cells workbook to CSV with a semicolon delimiter using C#
// Common Searches: c# aspnet export workbook to csv with semicolon separator using aspose.cells | how to set a custom column delimiter when saving a workbook as CSV in Aspose.Cells C# | asp.net core save Excel file as CSV with semicolon delimiter using TxtSaveOptions | example of TxtSaveOptions.Separator property for CSV output in C# | ensure output folder exists before calling Workbook.Save in Aspose.Cells
// Tags: Aspose.Cells CSV export custom separator | TxtSaveOptions separator property C# | Workbook.Save CSV semicolon delimiter | Create output directory before saving Aspose.Cells | C# Excel to CSV conversion with custom delimiter

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, adds sample data, configures TxtSaveOptions for CSV format with a semicolon as the column separator, ensures the target directory exists, and saves the workbook to 'output.csv', handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Configure CSV save options to use semicolon as the column delimiter
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.CSV);
            saveOptions.Separator = ';';

            // Define output file path
            string outputPath = "output.csv";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a CSV file with the custom delimiter
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
