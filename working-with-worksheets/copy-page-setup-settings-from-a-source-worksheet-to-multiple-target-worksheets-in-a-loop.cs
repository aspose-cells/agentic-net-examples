using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add several worksheets
            Workbook workbook = new Workbook();
            // Ensure we have at least 4 sheets (1 source + 3 targets)
            workbook.Worksheets.Add("Target1");
            workbook.Worksheets.Add("Target2");
            workbook.Worksheets.Add("Target3");

            // Reference to the source worksheet (first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Configure some page‑setup settings on the source sheet
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            sourceSheet.PageSetup.Orientation = PageOrientationType.Landscape;
            sourceSheet.PageSetup.FitToPagesWide = 1;
            sourceSheet.PageSetup.FitToPagesTall = 0;
            sourceSheet.PageSetup.PrintCopies = 2;

            // Loop through the target worksheets (indexes 1 to Count‑1)
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                Worksheet targetSheet = workbook.Worksheets[i];

                // Copy the page‑setup settings from the source to the target
                // Using default CopyOptions (no special behavior required)
                targetSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

                // Optional: rename the target sheet to indicate it received the settings
                targetSheet.Name = $"CopiedFromSource_{i}";
            }

            // Save the workbook to verify the result
            workbook.Save("PageSetupCopyResult.xlsx");
        }
    }
}