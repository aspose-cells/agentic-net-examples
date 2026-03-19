using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetDuplicationAndRelocationDemo
    {
        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the default first worksheet and add some sample data
            Worksheet originalSheet = workbook.Worksheets[0];
            originalSheet.Name = "OriginalSheet";
            originalSheet.Cells["A1"].PutValue("Header");
            originalSheet.Cells["A2"].PutValue(10);
            originalSheet.Cells["A3"].PutValue(20);
            originalSheet.Cells["B2"].Formula = "=SUM(A2:A3)"; // formula to test reference preservation

            // Duplicate the first worksheet using AddCopy (copies data, formats, and formulas)
            int copiedIndex = workbook.Worksheets.AddCopy(0);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "CopiedSheet";

            // Relocate the copied worksheet to position 2 (third tab) using MoveTo
            // If there are fewer than 2 sheets, MoveTo will place it at the end automatically
            copiedSheet.MoveTo(2);

            // Swap the first two worksheets to demonstrate sheet order manipulation
            // After this operation, "CopiedSheet" will be at index 0 and "OriginalSheet" at index 1
            workbook.Worksheets.SwapSheet(0, 1);

            // Save the workbook (lifecycle: save)
            workbook.Save("WorksheetDuplicationAndRelocationDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetDuplicationAndRelocationDemo.Run();
        }
    }
}