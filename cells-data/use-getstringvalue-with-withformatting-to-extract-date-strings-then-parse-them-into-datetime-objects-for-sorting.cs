// Title: Read formatted dates with GetStringValue, convert to DateTime, and sort rows using Aspose.Cells for .NET
// Common Searches: how to retrieve formatted date strings from Excel cells using Aspose.Cells C# | convert Excel cell string values to DateTime objects in Aspose.Cells .NET | sort worksheet rows by parsed dates with Aspose.Cells in C# | write sorted DateTime values to another column using Aspose.Cells for .NET | use GetStringValue with WithFormatting to read dates in Aspose.Cells
// Tags: GetStringValue WithFormatting Aspose.Cells | convert Excel string to DateTime C# | order rows by date Aspose.Cells | output sorted dates to new column Aspose.Cells | read formatted cell values .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, reads each cell in column A using GetStringValue with WithFormatting to obtain the displayed date string, parses non‑empty strings into DateTime objects, stores the original row indices, sorts the dates, writes the sorted DateTime values to column B, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Store parsed dates together with their original row index
            List<(int Row, DateTime Date)> dateRows = new List<(int, DateTime)>();

            // Assume dates are in column A (index 0)
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = sheet.Cells[row, 0];

                // Get the cell's displayed string (formatted)
                string dateStr = cell.StringValue;
                if (string.IsNullOrWhiteSpace(dateStr))
                    continue; // Skip empty cells

                // Parse the formatted string into a DateTime object
                if (DateTime.TryParse(dateStr, out DateTime parsedDate))
                {
                    dateRows.Add((row, parsedDate));
                }
            }

            // Sort the list by the DateTime value
            dateRows.Sort((a, b) => a.Date.CompareTo(b.Date));

            // Write the sorted dates to column B (index 1) in the same order as sorted
            for (int i = 0; i < dateRows.Count; i++)
            {
                var (_, date) = dateRows[i];
                // Put the DateTime value; Aspose.Cells will format it according to the cell's style
                sheet.Cells[i, 1].PutValue(date);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
