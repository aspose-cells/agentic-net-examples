// Title: Convert an Excel workbook to CSV and standardize 10‑digit phone numbers using Aspose.Cells for .NET
// AI Prompts: Iterate through every cell, replace any 10‑digit string with the pattern (XXX) XXX‑XXXX, then save the workbook as CSV via Aspose.Cells in C#. | Strip non‑numeric characters from string cells, format detected phone numbers, and export the updated worksheet to a CSV file using Aspose.Cells.
// Common Searches: Aspose.Cells C# normalize phone number display when converting Excel to CSV | How to apply a custom cell formatter for phone numbers during Excel to CSV conversion with Aspose.Cells | C# iterate Excel cells and reformat ten‑digit numbers using Aspose.Cells before saving as CSV | Save an .xlsx workbook as CSV after standardizing phone number format with Aspose.Cells .NET
// Tags: Aspose.Cells phone number normalization | Excel to CSV conversion Aspose.Cells | custom cell formatter Aspose.Cells | regex phone number transformation Aspose.Cells | SaveFormat.Csv Aspose.Cells example

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsPhoneNumberFormatter
{
    // The sample loads an Excel workbook, scans each cell for 10‑digit strings, rewrites them into the (XXX) XXX‑XXXX format, and then saves the modified workbook as a CSV file using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Get the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a simple phone number formatter:
            // This example removes all non‑digit characters and formats as (XXX) XXX‑XXXX
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                for (int col = 0; col <= cells.MaxDataColumn; col++)
                {
                    Cell cell = cells[row, col];
                    if (cell.Type == CellValueType.IsString)
                    {
                        string text = cell.StringValue.Trim();

                        // Basic check: if the string contains at least 10 digits, treat it as a phone number
                        string digitsOnly = System.Text.RegularExpressions.Regex.Replace(text, @"\D", "");
                        if (digitsOnly.Length == 10)
                        {
                            // Format as (XXX) XXX‑XXXX
                            string formatted = $"({digitsOnly.Substring(0, 3)}) {digitsOnly.Substring(3, 3)}-{digitsOnly.Substring(6, 4)}";
                            cell.PutValue(formatted);
                        }
                    }
                }
            }

            // Save the workbook as CSV using the provided Save rule
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);

            // Alternatively, demonstrate the ConversionUtility method for conversion
            // ConversionUtility.Convert(sourcePath, csvPath);
        }
    }
}
