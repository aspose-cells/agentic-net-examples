using System;
using Aspose.Cells;

class ApplyStrikethroughToCompletedRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Sample data: column A contains task status
        cells["A1"].PutValue("Task");
        cells["B1"].PutValue("Status");
        cells["A2"].PutValue("Task 1");
        cells["B2"].PutValue("Completed");
        cells["A3"].PutValue("Task 2");
        cells["B3"].PutValue("In Progress");
        cells["A4"].PutValue("Task 3");
        cells["B4"].PutValue("Completed");

        // Create a style with strikethrough font effect
        Style strikeStyle = workbook.CreateStyle();
        strikeStyle.Font.IsStrikeout = true; // enable strikeout

        // Create a style flag indicating that the FontStrike property should be applied
        StyleFlag flag = new StyleFlag();
        flag.FontStrike = true;

        // Apply the style to rows where the status is "Completed"
        int lastRow = cells.MaxDataRow;
        for (int row = 1; row <= lastRow; row++) // start from 1 to skip header
        {
            if (cells[row, 1].StringValue.Equals("Completed", StringComparison.OrdinalIgnoreCase))
            {
                // Apply the style to the entire row
                cells.ApplyRowStyle(row, strikeStyle, flag);
            }
        }

        // Save the workbook
        workbook.Save("CompletedTasks.xlsx");
    }
}