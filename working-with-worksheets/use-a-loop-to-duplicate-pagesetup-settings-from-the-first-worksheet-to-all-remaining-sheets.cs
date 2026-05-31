using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: set some page‑setup properties on the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            firstSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            firstSheet.PageSetup.Orientation = PageOrientationType.Landscape;
            firstSheet.PageSetup.PrintTitleRows = "$1:$1";
            firstSheet.PageSetup.PrintTitleColumns = "$A:$A";

            // Loop through all remaining worksheets and copy the page‑setup settings
            for (int i = 1; i < workbook.Worksheets.Count; i++)
            {
                Worksheet targetSheet = workbook.Worksheets[i];
                // Copy settings from the first worksheet using default CopyOptions
                targetSheet.PageSetup.Copy(firstSheet.PageSetup, new CopyOptions());
            }

            // Save the workbook to a file
            workbook.Save("PageSetupCopied.xlsx");
        }
    }
}