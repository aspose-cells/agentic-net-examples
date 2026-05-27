using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupClone
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains the source worksheet with the desired page setup
            Workbook templateWorkbook = new Workbook("template.xlsx");

            // Access the template worksheet (e.g., the first worksheet)
            Worksheet templateSheet = templateWorkbook.Worksheets[0];

            // Create a new workbook (or use an existing one) where the new sheet will be added
            Workbook destWorkbook = new Workbook();

            // Add a new blank worksheet to the destination workbook
            int newSheetIndex = destWorkbook.Worksheets.Add();
            Worksheet newSheet = destWorkbook.Worksheets[newSheetIndex];
            newSheet.Name = "ClonedPageSetup";

            // Clone the page setup from the template worksheet to the newly created worksheet
            // Use the Copy method of PageSetup with default CopyOptions
            newSheet.PageSetup.Copy(templateSheet.PageSetup, new CopyOptions());

            // (Optional) Verify that a property was copied, e.g., PaperSize
            Console.WriteLine("Cloned PaperSize: " + newSheet.PageSetup.PaperSize);

            // Save the destination workbook with the cloned page setup
            destWorkbook.Save("output.xlsx");
        }
    }
}