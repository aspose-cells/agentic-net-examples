using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample date strings (as text) in column A
            worksheet.Cells["A1"].PutValue("Date");               // Header
            worksheet.Cells["A2"].PutValue("15/05/2023");         // dd/MM/yyyy
            worksheet.Cells["A3"].PutValue("2023-04-20");         // yyyy-MM-dd
            worksheet.Cells["A4"].PutValue("01-06-2023");         // dd-MM-yyyy
            worksheet.Cells["A5"].PutValue("2023/03/10");         // yyyy/MM/dd

            // List to hold row index together with the parsed DateTime value
            List<(int Row, DateTime Value)> dateRows = new List<(int, DateTime)>();

            // Iterate over the data rows (skip header at row 0)
            for (int row = 1; row <= worksheet.Cells.MaxDataRow; row++)
            {
                Cell cell = worksheet.Cells[row, 0]; // Column A

                // Get the cell's displayed string (includes formatting)
                string dateString = cell.StringValue;

                // Try to parse the string into a DateTime object
                if (DateTime.TryParse(dateString, out DateTime parsedDate))
                {
                    dateRows.Add((row, parsedDate));
                }
            }

            // Sort the collected rows by the DateTime value (ascending)
            dateRows.Sort((x, y) => x.Value.CompareTo(y.Value));

            // Write the sorted dates back to the worksheet, starting at row 2
            for (int i = 0; i < dateRows.Count; i++)
            {
                int targetRow = i + 1; // Row index where the sorted date will be placed
                Cell targetCell = worksheet.Cells[targetRow, 0];

                // Store the DateTime value (Aspose.Cells will keep it as a date)
                targetCell.PutValue(dateRows[i].Value);

                // Apply a consistent date format for display
                Style style = targetCell.GetStyle();
                style.Custom = "yyyy-MM-dd";
                targetCell.SetStyle(style);
            }

            // Save the workbook with the sorted dates
            string outputPath = "SortedDates.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}