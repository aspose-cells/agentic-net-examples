using System;
using Aspose.Cells;

namespace LoadVisibleSheetsDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook with one hidden and one visible sheet
            // -----------------------------------------------------------------
            Workbook sourceWorkbook = new Workbook();

            // Add a visible sheet first
            Worksheet visibleSheet = sourceWorkbook.Worksheets.Add("VisibleSheet");
            visibleSheet.Cells["A1"].PutValue("Data in visible sheet");

            // Hide the default first sheet
            Worksheet hiddenSheet = sourceWorkbook.Worksheets[0];
            hiddenSheet.Name = "HiddenSheet";
            hiddenSheet.IsVisible = false;
            hiddenSheet.Cells["A1"].PutValue("Data in hidden sheet");

            // Save the workbook to a temporary file
            string filePath = "VisibleSheetsTest.xlsx";
            sourceWorkbook.Save(filePath);
            sourceWorkbook.Dispose();

            // ---------------------------------------------------------------
            // 2. Load the workbook (all sheets) and then keep only visible sheets
            // ---------------------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);

            // Remove hidden worksheets
            for (int i = loadedWorkbook.Worksheets.Count - 1; i >= 0; i--)
            {
                if (!loadedWorkbook.Worksheets[i].IsVisible)
                {
                    loadedWorkbook.Worksheets.RemoveAt(i);
                }
            }

            // ---------------------------------------------------------------
            // 3. Verify that only visible worksheets were retained
            // ---------------------------------------------------------------
            Console.WriteLine("Loaded Worksheets:");
            foreach (Worksheet ws in loadedWorkbook.Worksheets)
            {
                Console.WriteLine($"- {ws.Name} (Visible = {ws.IsVisible})");
                Console.WriteLine($"  A1 Value: {ws.Cells["A1"].StringValue}");
            }

            loadedWorkbook.Dispose();
        }
    }
}