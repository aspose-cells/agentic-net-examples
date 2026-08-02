using System;
using Aspose.Cells;

namespace AsposeCellsCopyStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Load the workbook template (lifecycle rule: use provided load constructor)
            Workbook workbook = new Workbook("Template.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Example data to be added as new records
            string[,] newData = new string[,]
            {
                { "John", "Doe", "30" },
                { "Jane", "Smith", "25" },
                { "Bob", "Johnson", "40" }
            };

            // Starting row index where new records will be inserted (0‑based)
            // Assuming the template has a header in row 0 and existing data starts at row 1
            int insertRowIndex = 1;

            // Loop through each record
            for (int i = 0; i < newData.GetLength(0); i++)
            {
                // Insert a new row and inherit formatting from the row above
                InsertOptions insertOptions = new InsertOptions();
                insertOptions.CopyFormatType = CopyFormatType.SameAsAbove; // CopyStyle attribute equivalent
                cells.InsertRows(insertRowIndex, 1, insertOptions);

                // Populate the newly inserted row with data
                for (int j = 0; j < newData.GetLength(1); j++)
                {
                    cells[insertRowIndex, j].PutValue(newData[i, j]);
                }

                // Move the insertion point down for the next record
                insertRowIndex++;
            }

            // Save the workbook (lifecycle rule: use provided save method)
            workbook.Save("OutputWithInheritedStyle.xlsx");
        }
    }
}