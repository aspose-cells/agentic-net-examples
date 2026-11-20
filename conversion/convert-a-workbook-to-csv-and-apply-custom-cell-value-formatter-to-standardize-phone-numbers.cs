// Title: Convert an Excel workbook to CSV and reformat US phone numbers using Aspose.Cells for .NET
// Common Searches: how to export each worksheet to a separate CSV file with Aspose.Cells C# | standardize phone number format (###) ###-#### during Excel to CSV conversion in .NET | apply regex to cell values before saving workbook as CSV using Aspose.Cells | save workbook as UTF-8 CSV with custom cell formatting using TxtSaveOptions in C# | replace non-digit characters and format 10-digit phone numbers in Excel with Aspose.Cells
// Tags: Aspose.Cells workbook to CSV conversion | custom cell value formatting Aspose.Cells | phone number normalization Excel C# | TxtSaveOptions UTF-8 CSV export | regex based cell transformation Aspose.Cells

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, scans all worksheets for string cells containing exactly ten digits, reformats those values to the (###) ###-#### pattern, and then saves each worksheet as a UTF‑8 CSV file using TxtSaveOptions.
class Program
{
    static void Main()
    {
        // Define input and output paths
        string inputPath = @"C:\Input\Workbook.xlsx";
        string outputPath = @"C:\Output\Worksheet1.csv";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Regular expression to match exactly 10 digits
            Regex phoneRegex = new Regex(@"\d{10}");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the used range of the worksheet
                Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                // Loop through each cell in the used range
                for (int row = usedRange.FirstRow; row <= usedRange.FirstRow + usedRange.RowCount - 1; row++)
                {
                    for (int col = usedRange.FirstColumn; col <= usedRange.FirstColumn + usedRange.ColumnCount - 1; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Process only string values that could represent phone numbers
                        if (cell.Type == CellValueType.IsString)
                        {
                            string original = cell.StringValue;

                            // Remove all non-digit characters
                            string digitsOnly = Regex.Replace(original, @"\D", "");

                            // If we have exactly 10 digits, format them as (###) ###-####
                            if (phoneRegex.IsMatch(digitsOnly))
                            {
                                string formatted = $"({digitsOnly.Substring(0, 3)}) {digitsOnly.Substring(3, 3)}-{digitsOnly.Substring(6, 4)}";
                                cell.PutValue(formatted);
                            }
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as CSV (each worksheet will be saved as a separate CSV file)
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8
            };
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine("Workbook converted to CSV with phone numbers standardized.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
