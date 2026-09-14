// Title: C# example: extract date strings with GetStringValue, convert to DateTime, apply a uniform format, and sort the column using Aspose.Cells
// AI Prompts: Generate C# code that reads each cell in a column using Aspose.Cells GetStringValue, attempts to parse the text with several date patterns into a DateTime, replaces the cell value with the DateTime, applies a yyyy‑MM‑dd number format to the range, and sorts the rows in ascending order. | Enhance the sample to recognize additional patterns such as "MM-dd-yyyy" and "dd MMM yyyy", then sort the dates in descending order while keeping the header row intact.
// Common Searches: Aspose.Cells C# how to convert mixed date strings to DateTime and sort the column | Using GetStringValue to read dates from Excel and sort with DataSorter in Aspose.Cells | Apply a consistent date format to a range after parsing string dates in Aspose.Cells .NET | Parse multiple date formats in an Excel worksheet with Aspose.Cells and C# | Sort Excel rows by date after converting string values using Aspose.Cells
// Tags: Aspose.Cells GetStringValue date extraction | Aspose.Cells DateTime conversion from string | Aspose.Cells uniform date format styling | Aspose.Cells DataSorter sort by date | Aspose.Cells parse multiple date patterns

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Demonstrates reading date strings from cells with GetStringValue, parsing them into DateTime using TryParseExact with several patterns, replacing the cell values, applying a uniform yyyy‑MM‑dd number format to the range, and sorting the column with DataSorter in Aspose.Cells for .NET.
public class DateStringSortingDemo
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }

    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate column A with header and date strings in various formats
            cells["A1"].PutValue("Date");
            cells["A2"].PutValue("2023-05-15");   // ISO format
            cells["A3"].PutValue("15/04/2023");   // European format
            cells["A4"].PutValue("2023/06/01");   // Slash format
            cells["A5"].PutValue("01-Jul-2023");  // Textual month format

            // Convert each string to a DateTime value
            for (int row = 1; row <= 5; row++) // rows 2 to 6 (0‑based index)
            {
                Cell cell = cells[row, 0];
                string dateStr = cell.StringValue; // get the displayed string
                if (string.IsNullOrWhiteSpace(dateStr))
                    continue;

                DateTime dt;
                // Define possible date patterns
                string[] patterns = {
                    "yyyy-MM-dd",
                    "dd/MM/yyyy",
                    "yyyy/MM/dd",
                    "dd-MMM-yyyy"
                };

                // Try exact parsing first, then fallback to general parsing
                bool parsed = DateTime.TryParseExact(dateStr, patterns,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out dt) ||
                    DateTime.TryParse(dateStr, out dt);

                if (parsed)
                {
                    // Replace the cell content with an actual DateTime value
                    cell.PutValue(dt);
                }
            }

            // Apply a uniform date display format to the column
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "yyyy-MM-dd";
            StyleFlag flag = new StyleFlag { NumberFormat = true };
            AsposeRange dateRange = cells.CreateRange("A2:A5");
            dateRange.ApplyStyle(dateStyle, flag);

            // Define the area to sort (including header)
            CellArea sortArea = CellArea.CreateCellArea("A1", "A5");

            // Configure the DataSorter to sort by the first column (index 0) ascending
            DataSorter sorter = workbook.DataSorter;
            sorter.AddKey(0, SortOrder.Ascending);
            sorter.Sort(cells, sortArea);

            // Save the workbook with sorted dates
            string outputPath = "SortedDates.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during processing: {ex.Message}");
        }
    }
}
