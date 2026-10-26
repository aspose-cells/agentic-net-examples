// Title: Apply a custom 'dd-MMM-yyyy' date format to a column in an Aspose.Cells table using C#
// Common Searches: C# Aspose.Cells set custom date format for a ListObject column | how to format table column dates as dd-MMM-yyyy with Aspose.Cells .NET | apply number format to specific column of an Excel table using Aspose.Cells | Aspose.Cells style a table column with custom date pattern in C# | create and format date column in Aspose.Cells ListObject example
// Tags: custom date format ListObject C# | apply number format to table column Aspose.Cells | style range for date cells Aspose.Cells | Excel table column formatting Aspose.Cells .NET | date style custom pattern dd-MMM-yyyy Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, adds a ListObject (table) with date values, defines a style using the custom format 'dd-MMM-yyyy', builds a range that targets the table's first data column, applies the style with a StyleFlag, and saves the file as FormattedDates.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var ws = workbook.Worksheets[0];

            // Add header and sample dates in column A
            ws.Cells["A1"].PutValue("Date");
            ws.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
            ws.Cells["A3"].PutValue(new DateTime(2023, 2, 20));
            ws.Cells["A4"].PutValue(new DateTime(2023, 3, 25));

            // Create a table (ListObject) covering A1:A4 with header
            int tableIndex = ws.ListObjects.Add(0, 0, 3, 0, true);
            var table = ws.ListObjects[tableIndex];
            table.DisplayName = "DateTable";

            // Define custom date format style
            var dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";

            // Apply the custom date format to the first column of the table (data rows only)
            // Build a range that represents the first data column of the table
            int startRow = table.DataRange.FirstRow;      // first data row (excluding header)
            int startColumn = table.DataRange.FirstColumn; // first column of the table
            int rowCount = table.DataRange.RowCount;      // number of data rows
            var columnRange = ws.Cells.CreateRange(startRow, startColumn, rowCount, 1);

            var styleFlag = new StyleFlag { NumberFormat = true };
            columnRange.ApplyStyle(dateStyle, styleFlag);

            // Ensure output directory exists
            string outputPath = "FormattedDates.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
