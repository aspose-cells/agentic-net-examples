// Title: Convert Excel to CSV and Standardize Phone Numbers with Aspose.Cells (C#)
// Description: Loads an .xlsx workbook, iterates through every used cell, detects phone‑number patterns with a regular expression, rewrites them to (123) 456‑7890, and saves the updated data as a CSV file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | .NET | Excel to CSV conversion | phone number formatting | regex cell processing | custom cell value transformation | workbook export | data cleansing | CSV generation
// Common Searches: Aspose.Cells convert workbook to CSV C# | reformat phone numbers in Excel before CSV export | apply regex to Excel cells using Aspose.Cells | standardize phone number format during Excel to CSV conversion | C# iterate over cells and modify values with Aspose.Cells
// Developer Intent: Load an Excel file, normalize any phone‑number strings to a consistent format, and export the cleaned workbook as a CSV document.
// Use Cases: Prepare contact lists for systems that require a uniform phone format. | Cleanse data before importing into a CRM or marketing platform. | Generate CSV reports where all telephone entries follow the (123) 456‑7890 pattern.
// AI Prompts: Show how to extend the regex to support international phone numbers while exporting to CSV. | Provide an example of implementing ICellValueFormatter in Aspose.Cells to format phone numbers during the save operation. | Suggest a method to log each cell that was modified from its original phone format.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsPhoneNumberFormatter
{
    // Loads an .xlsx workbook, iterates through every used cell, detects phone‑number patterns with a regular expression, rewrites them to (123) 456‑7890, and saves the updated data as a CSV file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Paths for source workbook and destination CSV
            string sourcePath = "input.xlsx";
            string csvPath = "output.csv";

            // Load the existing workbook (lifecycle: create & load)
            Workbook workbook = new Workbook(sourcePath);
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Regex to capture various phone number patterns and reformat to (123) 456-7890
            Regex phoneRegex = new Regex(@"\D*(\d{3})\D*(\d{3})\D*(\d{4})\D*");

            // Iterate through all cells in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only string cells
                    if (cell.Type == CellValueType.IsString)
                    {
                        string original = cell.StringValue;
                        Match match = phoneRegex.Match(original);

                        // If a phone number is detected, replace with standardized format
                        if (match.Success)
                        {
                            string formatted = $"({match.Groups[1].Value}) {match.Groups[2].Value}-{match.Groups[3].Value}";
                            cell.PutValue(formatted);
                        }
                    }
                }
            }

            // Save the workbook as CSV (using provided Save method)
            workbook.Save(csvPath, SaveFormat.Csv);
        }
    }
}
