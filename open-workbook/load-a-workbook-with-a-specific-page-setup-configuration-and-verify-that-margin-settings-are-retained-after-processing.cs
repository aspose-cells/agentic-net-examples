using System;
using Aspose.Cells;

namespace PageSetupMarginVerification
{
    class Program
    {
        static void Main()
        {
            // Define expected margin values (in centimeters)
            double expectedLeft = 1.0;
            double expectedRight = 1.5;
            double expectedTop = 2.0;
            double expectedBottom = 0.5;

            // Create a new workbook and set page margins
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.LeftMargin = expectedLeft;
            sheet.PageSetup.RightMargin = expectedRight;
            sheet.PageSetup.TopMargin = expectedTop;
            sheet.PageSetup.BottomMargin = expectedBottom;

            // Save the workbook to a file
            string filePath = "MarginDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Load the workbook from the file
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            PageSetup loadedSetup = loadedSheet.PageSetup;

            // Verify that the margins are retained
            bool leftOk = Math.Abs(loadedSetup.LeftMargin - expectedLeft) < 0.0001;
            bool rightOk = Math.Abs(loadedSetup.RightMargin - expectedRight) < 0.0001;
            bool topOk = Math.Abs(loadedSetup.TopMargin - expectedTop) < 0.0001;
            bool bottomOk = Math.Abs(loadedSetup.BottomMargin - expectedBottom) < 0.0001;

            Console.WriteLine("Margin verification results:");
            Console.WriteLine($"LeftMargin: {(leftOk ? "OK" : "Mismatch")} (Expected: {expectedLeft}, Actual: {loadedSetup.LeftMargin})");
            Console.WriteLine($"RightMargin: {(rightOk ? "OK" : "Mismatch")} (Expected: {expectedRight}, Actual: {loadedSetup.RightMargin})");
            Console.WriteLine($"TopMargin: {(topOk ? "OK" : "Mismatch")} (Expected: {expectedTop}, Actual: {loadedSetup.TopMargin})");
            Console.WriteLine($"BottomMargin: {(bottomOk ? "OK" : "Mismatch")} (Expected: {expectedBottom}, Actual: {loadedSetup.BottomMargin})");
        }
    }
}