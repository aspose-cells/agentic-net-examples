using System;
using Aspose.Cells;

namespace AsposeCellsWorksheetDuplication
{
    public class DuplicateAndReorder
    {
        public static void Run()
        {
            // Create a new workbook and remove the default sheet
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();

            // Add sample worksheets
            Worksheet ws1 = workbook.Worksheets.Add("Sheet1");
            ws1.Cells["A1"].PutValue("Data in Sheet1");

            Worksheet ws2 = workbook.Worksheets.Add("Sheet2");
            ws2.Cells["A1"].PutValue("Data in Sheet2");

            Worksheet ws3 = workbook.Worksheets.Add("Sheet3");
            ws3.Cells["A1"].PutValue("Data in Sheet3");

            // Duplicate the worksheet named "Sheet2"
            int copiedIndex = workbook.Worksheets.AddCopy("Sheet2");
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "Sheet2_Copy";

            // Relocate the copied sheet to the first position (index 0)
            copiedSheet.MoveTo(0);

            // Save the workbook to verify the result
            workbook.Save("DuplicatedAndReordered.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DuplicateAndReorder.Run();
        }
    }
}