using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTabIdLogger
{
    class Program
    {
        // Logs changes of TabId values
        private static void SetTabId(Worksheet worksheet, int newTabId, List<string> log)
        {
            int originalTabId = worksheet.TabId;
            if (originalTabId != newTabId)
            {
                log.Add($"Worksheet \"{worksheet.Name}\": TabId changed from {originalTabId} to {newTabId}");
                worksheet.TabId = newTabId;
            }
        }

        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare a log collection
            List<string> tabIdChangeLog = new List<string>();

            // Access the first worksheet and change its TabId
            Worksheet sheet1 = workbook.Worksheets[0];
            SetTabId(sheet1, 101, tabIdChangeLog);

            // Add a second worksheet and change its TabId
            int sheetIndex = workbook.Worksheets.Add();
            Worksheet sheet2 = workbook.Worksheets[sheetIndex];
            SetTabId(sheet2, 202, tabIdChangeLog);

            // Change TabId again on the first worksheet
            SetTabId(sheet1, 303, tabIdChangeLog);

            // Output the log to console
            Console.WriteLine("TabId Change Log:");
            foreach (string entry in tabIdChangeLog)
            {
                Console.WriteLine(entry);
            }

            // Save the workbook
            string outputPath = "TabIdLoggerOutput.xlsx";
            workbook.Save(outputPath);
        }
    }
}