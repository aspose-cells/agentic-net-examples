using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetGridlinesAndHeadersDemo
    {
        public static void Run()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Enable gridlines visibility
            sheet.IsGridlinesVisible = true;

            // Disable row and column header labels (A, B, 1, 2, etc.)
            sheet.IsRowColumnHeadersVisible = false;

            // Add some sample data so the effect can be seen when opened
            sheet.Cells["A1"].PutValue("Gridlines: ON");
            sheet.Cells["A2"].PutValue("Headers: OFF");

            // Save the workbook
            workbook.Save("GridlinesOn_HeadersOff.xlsx");

            // Load the saved workbook to demonstrate that the settings persist
            Workbook loaded = new Workbook("GridlinesOn_HeadersOff.xlsx");
            Worksheet loadedSheet = loaded.Worksheets[0];

            // Output current settings to the console
            Console.WriteLine("IsGridlinesVisible: " + loadedSheet.IsGridlinesVisible);
            Console.WriteLine("IsRowColumnHeadersVisible: " + loadedSheet.IsRowColumnHeadersVisible);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetGridlinesAndHeadersDemo.Run();
        }
    }
}