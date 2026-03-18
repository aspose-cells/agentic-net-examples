using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveUnusedStylesDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet sheet = wb.Worksheets[0];

            // Populate cells with data and apply distinct styles
            for (int i = 0; i < 10; i++)
            {
                Cell cell = sheet.Cells[i, 0];
                cell.PutValue("Item " + (i + 1));

                // Create a style for the current cell
                Style style = wb.CreateStyle();
                style.Font.Name = "Arial";
                style.Font.Size = 10 + i;
                style.Font.IsBold = i % 2 == 0;

                // Apply the style to the cell
                cell.SetStyle(style);
            }

            // Delete rows to make some styles unused
            sheet.Cells.DeleteRows(5, 5);

            // Save before cleaning up unused styles
            wb.Save("BeforeRemoveUnusedStyles.xlsx");

            // Remove all styles that are no longer referenced in the workbook
            wb.RemoveUnusedStyles();

            // Save after removing unused styles
            wb.Save("AfterRemoveUnusedStyles.xlsx");
        }
    }
}