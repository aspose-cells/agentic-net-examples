using System;
using Aspose.Cells;

namespace AsposeCellsWarningDemo
{
    public class ConsoleWarningCallback : IWarningCallback
    {
        public void Warning(WarningInfo warningInfo)
        {
            Console.WriteLine($"Warning Type: {warningInfo.Type}");
            Console.WriteLine($"Description : {warningInfo.Description}");
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (no external file required)
            Workbook workbook = new Workbook();

            // Assign the warning callback to the workbook settings
            workbook.Settings.WarningCallback = new ConsoleWarningCallback();

            // Trigger a warning by adding duplicate defined names
            int idx1 = workbook.Worksheets.Names.Add("DuplicateName");
            workbook.Worksheets.Names[idx1].RefersTo = "=Sheet1!$A$1";

            int idx2 = workbook.Worksheets.Names.Add("DuplicateName");
            workbook.Worksheets.Names[idx2].RefersTo = "=Sheet1!$A$2";

            // Save the workbook; the duplicate name warning will be reported via the callback
            workbook.Save("OutputWithWarnings.xlsx");
        }
    }
}