using System;
using Aspose.Cells;

namespace AsposeCellsPageSetupCopyDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (source and destination are in the same workbook)
            Workbook workbook = new Workbook();

            // Access the source worksheet (first sheet) and configure its page setup
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            sourceSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;
            sourceSheet.PageSetup.Orientation = PageOrientationType.Landscape;
            sourceSheet.PageSetup.PrintArea = "A1:D20";
            sourceSheet.PageSetup.CenterHorizontally = true;
            sourceSheet.PageSetup.CenterVertically = true;

            // Add a destination worksheet
            Worksheet destSheet = workbook.Worksheets.Add("DestinationSheet");

            // Copy the entire page‑setup configuration from source to destination
            // Using PageSetup.Copy with default CopyOptions
            destSheet.PageSetup.Copy(sourceSheet.PageSetup, new CopyOptions());

            // Verify the copy (optional)
            Console.WriteLine("Destination sheet paper size: " + destSheet.PageSetup.PaperSize);
            Console.WriteLine("Destination sheet orientation: " + destSheet.PageSetup.Orientation);
            Console.WriteLine("Destination sheet print area: " + destSheet.PageSetup.PrintArea);
            Console.WriteLine("Center horizontally: " + destSheet.PageSetup.CenterHorizontally);
            Console.WriteLine("Center vertically: " + destSheet.PageSetup.CenterVertically);

            // Save the workbook to a file
            workbook.Save("PageSetupCopyResult.xlsx");
        }
    }
}