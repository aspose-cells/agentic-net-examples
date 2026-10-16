// Title: Save an Aspose.Cells workbook as a UTF-8 CSV file with a Byte Order Mark in C#
// Common Searches: how to export a workbook to CSV with UTF-8 BOM using Aspose.Cells in C# | Aspose.Cells C# save worksheet as CSV with byte order mark | set UTF8Encoding(true) for CSV output in Aspose.Cells .NET | C# TxtSaveOptions CSV encoding with BOM Aspose.Cells example | write CSV file with BOM from Aspose.Cells workbook C#
// Tags: Aspose.Cells CSV export BOM | TxtSaveOptions CSV UTF-8 encoding | C# workbook to CSV with Byte Order Mark | UTF8Encoding true Aspose.Cells CSV | export workbook as CSV with BOM

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example creates a Workbook, fills it with sample data, configures TxtSaveOptions to use UTF8Encoding(true) so a Byte Order Mark is written, and saves the workbook as a CSV file named 'output.csv', handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate the first worksheet with sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(25);

            // Set CSV save options to use UTF-8 encoding with a Byte Order Mark (BOM)
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
            csvOptions.Encoding = new UTF8Encoding(true); // 'true' ensures the BOM is written

            // Define output file path
            string outputPath = "output.csv";

            // Export the workbook to a CSV file with the specified options
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook successfully saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
