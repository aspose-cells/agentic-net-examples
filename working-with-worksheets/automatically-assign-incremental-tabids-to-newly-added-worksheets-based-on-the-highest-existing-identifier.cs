using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class IncrementalTabIdDemo
    {
        public static void Run()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Add a few initial worksheets with custom TabIds to simulate existing identifiers
            Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
            sheet1.Name = "Sheet1";
            sheet1.TabId = 100; // set an initial TabId

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.TabId = 105; // another existing TabId

            // Function to assign the next incremental TabId to a worksheet
            void AssignNextTabId(Worksheet ws)
            {
                int maxTabId = 0;
                // Iterate through all existing worksheets to find the highest TabId
                foreach (Worksheet existing in workbook.Worksheets)
                {
                    if (existing.TabId > maxTabId)
                        maxTabId = existing.TabId;
                }
                // Set the new worksheet's TabId to max + 1
                ws.TabId = maxTabId + 1;
            }

            // Add a new worksheet and assign incremental TabId
            Worksheet newSheet = workbook.Worksheets.Add("NewSheet");
            AssignNextTabId(newSheet);
            Console.WriteLine($"Added '{newSheet.Name}' with TabId: {newSheet.TabId}");

            // Add another worksheet to demonstrate continued increment
            Worksheet anotherSheet = workbook.Worksheets.Add("AnotherSheet");
            AssignNextTabId(anotherSheet);
            Console.WriteLine($"Added '{anotherSheet.Name}' with TabId: {anotherSheet.TabId}");

            // Save the workbook (save rule)
            workbook.Save("IncrementalTabIdDemo.xlsx");
        }
    }
}