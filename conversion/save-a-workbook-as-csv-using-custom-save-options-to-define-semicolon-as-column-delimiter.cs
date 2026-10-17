// Title: How to save an Aspose.Cells workbook as a semicolon‑delimited CSV file using TxtSaveOptions in C#
// Common Searches: C# Aspose.Cells export worksheet to CSV with semicolon separator | Set custom column delimiter when saving workbook to CSV using Aspose.Cells TxtSaveOptions | Aspose.Cells TxtSaveOptions Separator property example in C# | Generate semicolon delimited CSV from Aspose.Cells workbook programmatically | Save workbook as CSV with custom delimiter using Aspose.Cells .NET
// Tags: Aspose.Cells TxtSaveOptions CSV delimiter | C# semicolon delimited CSV export | Aspose.Cells custom column separator | Save workbook as CSV Aspose.Cells | TxtSaveOptions Separator property C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example creates an Aspose.Cells workbook, fills it with sample data, configures TxtSaveOptions to use a semicolon as the column separator for CSV output, and saves the workbook to a CSV file, demonstrating error handling in C#.
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

            // Set up TXT save options to use a semicolon as the column delimiter for CSV output
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.CSV);
            txtOptions.Separator = ';';

            // Define output path
            string outputPath = "output.csv";

            // Save the workbook as a CSV file with the custom delimiter
            workbook.Save(outputPath, txtOptions);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
