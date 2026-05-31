using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;   // Required for ListObject

class SetTableColumnDateFormat
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a header and a few dates
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 15));
            worksheet.Cells["A3"].PutValue(new DateTime(2023, 2, 20));
            worksheet.Cells["A4"].PutValue(new DateTime(2023, 3, 25));

            // Create a table (ListObject) that includes the header and data
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = worksheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "DateTable";

            // Create a style with the desired custom date format
            Style dateStyle = workbook.CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";

            // Apply the style to the entire column of the table (the date column)
            int dateColumnIndex = table.StartColumn; // column index of the first (date) column

            // Apply to header cell
            worksheet.Cells[table.StartRow, dateColumnIndex].SetStyle(dateStyle);

            // Apply to each data cell in the column
            if (table.DataRange != null)
            {
                int firstRow = table.DataRange.FirstRow;
                int lastRow = firstRow + table.DataRange.RowCount - 1;

                for (int row = firstRow; row <= lastRow; row++)
                {
                    worksheet.Cells[row, dateColumnIndex].SetStyle(dateStyle);
                }
            }

            // Save the workbook (ensure the directory exists)
            string outputPath = "TableDateFormat.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}