using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class FreezePanesDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Freeze panes at cell C3 (row index 2, column index 2) with 3 frozen rows and 3 frozen columns
            worksheet.FreezePanes(2, 2, 3, 3);

            // Verify that panes are frozen
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = worksheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);
            Console.WriteLine($"Freeze applied: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {frozenRow}, Column: {frozenColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // Save the workbook to a file
            workbook.Save("FreezePanesDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            FreezePanesDemo.Run();
        }
    }
}