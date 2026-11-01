// Title: Extract raw date strings from column A in an Excel workbook using Aspose.Cells for .NET, convert them to UTC ISO‑8601, and write the results to column B
// AI Prompts: Load an .xlsx file with Aspose.Cells, read the StringValue of each cell in column A, parse it to a DateTime, format it as UTC ISO‑8601, and store the formatted string in column B. | Write a C# loop that iterates over all used rows, converts the raw date text from column A to the pattern "yyyy-MM-ddTHH:mm:ssZ" using CultureInfo.InvariantCulture, and saves the workbook with the new values in the adjacent column. | Create a method that accepts an input Excel path, transforms every date string in the first column to an ISO‑8601 UTC string, writes the output to the next column, and writes the result to a specified output file.
// Common Searches: Aspose.Cells .NET how to read cell text as string and convert to ISO 8601 UTC | C# iterate over Excel rows and change date format with Aspose.Cells | Convert Excel column of date strings to ISO 8601 using Aspose.Cells and write to another column | Parse raw date string from Excel cell with CultureInfo.InvariantCulture in Aspose.Cells | Save modified workbook after updating date format in Aspose.Cells C#
// Tags: raw string extraction Aspose.Cells | date string parsing CultureInfo Aspose.Cells | ISO‑8601 formatting C# Aspose.Cells | adjacent column write Aspose.Cells | workbook save after cell update Aspose.Cells

using System;
using System.Globalization;
using Aspose.Cells;

// // Loads an Excel workbook, reads raw string values from column A, parses each as a DateTime, converts to UTC ISO‑8601 format, writes the formatted strings to column B, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the last used row in the sheet
        int maxRow = sheet.Cells.MaxDataRow;

        // Loop through each used row
        for (int row = 0; row <= maxRow; row++)
        {
            // Assume the date is stored in column A (index 0)
            Cell dateCell = sheet.Cells[row, 0];

            // Retrieve the raw string value exactly as it appears in the cell
            string rawDateString = dateCell.StringValue;

            // Skip empty cells
            if (string.IsNullOrWhiteSpace(rawDateString))
                continue;

            // Try to parse the raw string into a DateTime object
            if (DateTime.TryParse(rawDateString, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                // Convert the DateTime to ISO 8601 format (UTC)
                string iso8601String = parsedDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");

                // Write the ISO string to column B (index 1) of the same row
                sheet.Cells[row, 1].PutValue(iso8601String);
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
