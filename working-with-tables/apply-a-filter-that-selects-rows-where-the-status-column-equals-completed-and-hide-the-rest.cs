using System;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data (Header + rows) -----
            // Header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Task");
            sheet.Cells["C1"].PutValue("Status");

            // Data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Design");
            sheet.Cells["C2"].PutValue("Completed");

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Development");
            sheet.Cells["C3"].PutValue("In Progress");

            sheet.Cells["A4"].PutValue(4);
            sheet.Cells["B4"].PutValue("Testing");
            sheet.Cells["C4"].PutValue("Completed");

            sheet.Cells["A5"].PutValue(5);
            sheet.Cells["B5"].PutValue("Deployment");
            sheet.Cells["C5"].PutValue("Pending");

            // ----- Apply AutoFilter -----
            // Define the range that contains the header and data rows
            sheet.AutoFilter.Range = "A1:C5";

            // Filter the 'Status' column (index 2, i.e., column C) for the value "Completed"
            sheet.AutoFilter.Filter(2, "Completed");

            // Refresh the filter to hide rows that do not meet the criteria
            sheet.AutoFilter.Refresh();

            // Optional: verify which rows are hidden
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                bool hidden = sheet.Cells.IsRowHidden(row);
                Console.WriteLine($"Row {row + 1} hidden: {hidden}");
            }

            // Save the workbook
            workbook.Save("FilteredByStatus.xlsx", SaveFormat.Xlsx);
        }
    }
}