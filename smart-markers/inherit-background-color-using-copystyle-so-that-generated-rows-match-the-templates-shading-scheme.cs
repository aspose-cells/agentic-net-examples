using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsRowStyleCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the desired row shading
            Workbook workbook = new Workbook("Template.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];

            // Source row whose background color (and other settings) will be copied
            Row sourceRow = worksheet.Cells.Rows[0]; // assume first row has the template shading

            // Insert a new blank row at the desired position (e.g., row index 5)
            int insertIndex = 5;
            worksheet.Cells.InsertRows(insertIndex, 1);

            // Target row that was just inserted
            Row targetRow = worksheet.Cells.Rows[insertIndex];

            // Copy all settings from the source row, including style (background color)
            // The second parameter true ensures style is gathered when workbooks differ
            targetRow.CopySettings(sourceRow, true);

            // Optionally, verify that the background color was inherited
            Style targetStyle = targetRow.GetStyle();
            Console.WriteLine("Target row background color: " + targetStyle.BackgroundColor.Name);

            // Save the modified workbook
            workbook.Save("Result.xlsx");
        }
    }
}