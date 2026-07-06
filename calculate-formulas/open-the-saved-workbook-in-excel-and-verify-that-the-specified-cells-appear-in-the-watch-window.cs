using System;
using Aspose.Cells;

namespace AsposeCellsWatchDemo
{
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook and add cell watches ----------
            Workbook workbook = new Workbook();                     // create workbook (rule: Workbook())
            Worksheet sheet = workbook.Worksheets[0];              // get first worksheet

            // Add sample data (optional, just to have content)
            sheet.Cells["B2"].PutValue("First Watch");
            sheet.Cells["C3"].PutValue("Second Watch");

            // Add cell watches for B2 and C3 using the Add(string) method
            int indexB2 = sheet.CellWatches.Add("B2");             // add watch for B2
            int indexC3 = sheet.CellWatches.Add("C3");             // add watch for C3

            // Save the workbook to disk (rule: Save(string))
            string filePath = "CellWatchDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Load the saved workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath);       // load workbook (rule: Workbook(string))
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0]; // get first worksheet

            // Verify that the watches exist
            bool hasB2 = false;
            bool hasC3 = false;

            // Iterate through the CellWatchCollection
            foreach (CellWatch watch in loadedSheet.CellWatches)
            {
                if (watch.CellName.Equals("B2", StringComparison.OrdinalIgnoreCase))
                    hasB2 = true;
                if (watch.CellName.Equals("C3", StringComparison.OrdinalIgnoreCase))
                    hasC3 = true;
            }

            // Output verification results
            Console.WriteLine($"Watch for B2 present: {hasB2}");
            Console.WriteLine($"Watch for C3 present: {hasC3}");

            // Optionally, demonstrate accessing a watch by index
            CellWatch watchB2 = loadedSheet.CellWatches[indexB2];
            Console.WriteLine($"Accessed by index {indexB2}: CellName={watchB2.CellName}, Row={watchB2.Row}, Column={watchB2.Column}");
        }
    }
}