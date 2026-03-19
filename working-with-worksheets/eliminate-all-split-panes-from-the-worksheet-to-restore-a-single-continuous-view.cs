using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveAllSplitPanes
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (or any specific worksheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Remove any split window if present
            worksheet.RemoveSplit();

            // Unfreeze panes in case they were frozen (optional but ensures a single view)
            worksheet.UnFreezePanes();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveAllSplitPanes.Run();
        }
    }
}