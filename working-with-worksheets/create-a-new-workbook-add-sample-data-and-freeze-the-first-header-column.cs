using System;
using Aspose.Cells;

namespace AsposeCellsFreezeFirstColumnDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample header data
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["C1"].PutValue("Score");

            // Add some sample rows
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["C2"].PutValue(85);

            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");
            worksheet.Cells["C3"].PutValue(92);

            worksheet.Cells["A4"].PutValue(3);
            worksheet.Cells["B4"].PutValue("Charlie");
            worksheet.Cells["C4"].PutValue(78);

            // Freeze the first header column (column A)
            // Freeze at cell B1 with 0 frozen rows and 1 frozen column
            worksheet.FreezePanes("B1", 0, 1);

            // Save the workbook to a file
            workbook.Save("FreezeFirstHeaderColumn.xlsx");
        }
    }
}