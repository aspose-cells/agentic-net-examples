// Title: Save an Aspose.Cells workbook as a UTF-32 encoded CSV file using TxtSaveOptions in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, populates it with data, and saves it as a CSV file using TxtSaveOptions with UTF-32 encoding. | Show how to configure TxtSaveOptions for CSV export with Encoding.UTF32 to handle large datasets in Aspose.Cells. | Demonstrate setting the Encoding property of TxtSaveOptions before calling Workbook.Save to produce a UTF-32 CSV file.
// Common Searches: Aspose.Cells C# export workbook to CSV with UTF-32 encoding | How to set UTF-32 encoding for CSV output using TxtSaveOptions in Aspose.Cells | Saving large Excel data as UTF-32 CSV with Aspose.Cells .NET | TxtSaveOptions SaveFormat.Csv custom encoding example | C# Aspose.Cells CSV export encoding options
// Tags: TxtSaveOptions CSV UTF-32 encoding | Aspose.Cells workbook to CSV conversion | C# export large dataset as UTF-32 CSV | Save workbook with custom encoding Aspose.Cells | CSV export encoding Aspose.Cells .NET

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // The example creates a workbook, fills it with sample rows, configures TxtSaveOptions for CSV format with UTF-32 encoding, ensures the output directory exists, and saves the workbook as a UTF-32 encoded CSV file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Populate the workbook with sample data (replace with your own data as needed)
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Example: fill first 10 rows with sample data
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue($"Row {i + 1}");
                    cells[i, 1].PutValue(i * 10);
                    cells[i, 2].PutValue(DateTime.Now.AddDays(i));
                }

                // Configure CSV save options with UTF-32 encoding using TxtSaveOptions
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    // Set the desired encoding to UTF-32
                    Encoding = Encoding.UTF32

                    // ConvertNumericData property is no longer available; default behavior is sufficient
                };

                // Define output path and ensure the directory exists
                string outputPath = "LargeDatasetExport.csv";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                // If the file is in the current directory, outputDir may be null
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a CSV file using the custom options
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved as CSV with UTF-32 encoding to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
