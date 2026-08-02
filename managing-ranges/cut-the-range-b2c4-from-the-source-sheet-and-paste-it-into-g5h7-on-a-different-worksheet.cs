// Title: Cut range B2:C4 from a source worksheet and insert it into G5:H7 on another sheet using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, populate cells B2:C4 on a source sheet, define the range, add a destination sheet, and use Aspose.Cells' InsertCutCells with ShiftType.Down to cut the range and paste it starting at G5 (covering G5:H7) on the destination sheet. The workbook is then saved to a file.
// Keywords: Aspose.Cells cut range C# | InsertCutCells .NET | move cells between worksheets Aspose | cut and paste B2:C4 to G5 | ShiftType.Down Aspose.Cells | C# Excel range manipulation | Aspose.Cells example cut paste
// Common Searches: how to cut and paste a range between worksheets using Aspose.Cells | Aspose.Cells InsertCutCells example G5 | cut range B2:C4 and paste to G5:H7 Aspose | ShiftType options for InsertCutCells Aspose.Cells | C# code to move cells from one sheet to another Aspose
// Developer Intent: Cut the B2:C4 range from the source worksheet and insert it into cells G5:H7 on a different worksheet.
// Use Cases: Reorganize a data block by moving it from a raw data sheet to a formatted report sheet. | Automate the transfer of calculated results from a processing sheet to a summary sheet at a fixed position.
// AI Prompts: Write C# code with Aspose.Cells that cuts the range B2:C4 from one worksheet and inserts it into G5:H7 on another worksheet, including proper error handling. | Explain the impact of each ShiftType (Down, Right, Up, Left) when using InsertCutCells to cut and paste ranges in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsCutPasteDemo
{
    // Demonstrates how to create a workbook, populate cells B2:C4 on a source sheet, define the range, add a destination sheet, and use Aspose.Cells' InsertCutCells with ShiftType.Down to cut the range and paste it starting at G5 (covering G5:H7) on the destination sheet. The workbook is then saved to a file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Source worksheet (first sheet)
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "Source";

                // Destination worksheet (add a new sheet)
                Worksheet destSheet = workbook.Worksheets.Add("Destination");

                // Populate the source range B2:C4 with sample data
                sourceSheet.Cells["B2"].PutValue("Item1");
                sourceSheet.Cells["B3"].PutValue("Item2");
                sourceSheet.Cells["B4"].PutValue("Item3");
                sourceSheet.Cells["C2"].PutValue(10);
                sourceSheet.Cells["C3"].PutValue(20);
                sourceSheet.Cells["C4"].PutValue(30);

                // Create the range to cut (B2:C4) on the source sheet
                AsposeRange cutRange = sourceSheet.Cells.CreateRange("B2:C4");

                // Insert the cut range into the destination sheet at G5 (row index 4, column index 6)
                // ShiftType.Down is used; adjust if a different shift behavior is required.
                destSheet.Cells.InsertCutCells(cutRange, 4, 6, ShiftType.Down);

                // Define output file name
                string outputFile = "CutPasteDemo.xlsx";

                // Save the workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputFile)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
