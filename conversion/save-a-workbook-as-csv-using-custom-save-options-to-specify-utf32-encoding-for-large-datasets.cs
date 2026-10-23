// Title: Save a large Excel workbook to CSV with UTF‑32 encoding using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to CSV with UTF-32 encoding for large data sets | How to set UTF-32 encoding when saving CSV with Aspose.Cells TxtSaveOptions | C# save Excel worksheet as CSV using UTF-32 with Aspose.Cells | Generate CSV from a 10,000‑row Excel file with custom encoding in .NET | Aspose.Cells example for CSV conversion with UTF-32 encoding
// Tags: Aspose.Cells TxtSaveOptions CSV encoding | C# export workbook to CSV UTF-32 | large dataset CSV conversion Aspose.Cells | .NET custom CSV encoding Aspose.Cells | save Excel as CSV with UTF-32 using Aspose

using Aspose.Cells;
using System;
using System.Text;

// The example creates a new Workbook, populates the first worksheet with 10,000 rows of sample data, configures TxtSaveOptions for CSV output with UTF‑32 encoding and a comma separator, and saves the workbook as "LargeDataset.csv" while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with a large dataset (example)
            for (int row = 0; row < 10000; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
                sheet.Cells[row, 1].PutValue(row);
            }

            // Set up TXT save options for CSV output with UTF‑32 encoding
            TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF32,   // Specify UTF‑32 encoding
                Separator = ','              // Use comma as delimiter
                // Numeric data is saved as numbers by default
            };

            // Save the workbook as a CSV file using the custom options
            workbook.Save("LargeDataset.csv", txtOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
