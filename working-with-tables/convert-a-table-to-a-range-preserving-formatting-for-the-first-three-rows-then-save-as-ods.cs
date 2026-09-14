// Title: Convert an Excel ListObject to a regular range, keep formatting of the first three rows, and save as ODS with Aspose.Cells for .NET
// AI Prompts: Transform the first ListObject in a workbook into a plain range, copy the styles of the top three rows, and export the result to ODS using Aspose.Cells. | Extract a table from an XLSX file, preserve header and the first two data rows' styling, convert it to a range, then save the workbook as an ODS file in C#.
// Common Searches: Aspose.Cells C# convert ListObject to range while preserving row styles | how to keep first three rows formatting after removing Excel table using Aspose.Cells | save workbook as ODS after converting Excel table to range in .NET | C# code to copy styles from table rows before converting to range with Aspose.Cells | Aspose.Cells export to ODS format after table conversion
// Tags: listobject conversion to range Aspose.Cells | preserve top rows cell styles when converting table | save workbook as ODS using Aspose.Cells | clone cell style from table rows C# | remove Excel table while keeping header formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example loads an XLSX workbook, captures the styles of the first three rows of the first table, converts the ListObject to a regular range, reapplies the saved styles, and finally saves the workbook in ODS format.
class TableToRangeConverter
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.ods";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one table (ListObject) on the worksheet
            if (sheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found on the worksheet.");
                return;
            }

            // Get the first table
            ListObject table = sheet.ListObjects[0];

            // Determine the full range of the table (including header if present)
            // In Aspose.Cells the property is ShowHeaderRow
            bool hasHeader = table.ShowHeaderRow;
            int firstRow = table.DataRange.FirstRow - (hasHeader ? 1 : 0);
            int firstColumn = table.DataRange.FirstColumn;
            int rowCount = table.DataRange.RowCount + (hasHeader ? 1 : 0);
            int columnCount = table.DataRange.ColumnCount;

            // Create a Range object representing the table area
            Aspose.Cells.Range tableRange = sheet.Cells.CreateRange(firstRow, firstColumn, rowCount, columnCount);

            // Store the style of each cell in the first three rows of the table
            int firstRows = Math.Min(3, tableRange.RowCount);
            Style[,] savedStyles = new Style[firstRows, tableRange.ColumnCount];

            for (int r = 0; r < firstRows; r++)
            {
                for (int c = 0; c < tableRange.ColumnCount; c++)
                {
                    // Clone the style by creating a new Style and copying the original
                    Style original = tableRange[r, c].GetStyle();
                    Style copy = workbook.CreateStyle();
                    copy.Copy(original);
                    savedStyles[r, c] = copy;
                }
            }

            // Convert the table to a normal range (the ListObject is removed)
            table.ConvertToRange();

            // Reapply the saved styles to the first three rows
            for (int r = 0; r < firstRows; r++)
            {
                for (int c = 0; c < tableRange.ColumnCount; c++)
                {
                    Cell cell = tableRange[r, c];
                    cell.SetStyle(savedStyles[r, c]);
                }
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
