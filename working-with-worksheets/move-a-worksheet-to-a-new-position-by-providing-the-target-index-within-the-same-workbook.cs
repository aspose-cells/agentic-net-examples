// Title: Move a worksheet to a specific index in a workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add worksheets, and reposition a sheet using the Worksheet.MoveTo method (C#). The example moves "MovedSheet" to index 1 and saves the file as MovedWorksheetDemo.xlsx.
// Keywords: Aspose.Cells C# | .NET Worksheet.MoveTo | move worksheet index | reorder workbook sheets | change sheet order programmatically | Aspose.Cells API example
// Common Searches: Aspose.Cells move worksheet to index | C# Worksheet.MoveTo example | reorder worksheets in Aspose.Cells | change sheet order Aspose.Cells .NET | move sheet to first position C#
// Developer Intent: The developer needs to change the position of an existing worksheet within the same workbook by specifying a target index.
// Use Cases: Place a summary sheet at the beginning of the workbook (index 0) before distribution. | After generating a temporary analysis sheet, move it to a logical spot in the sheet order for clearer navigation. | Allow end‑users to customize workbook layout by moving selected worksheets to preferred positions at runtime.
// AI Prompts: Write C# code that uses Aspose.Cells to move a worksheet named "Report" to index 0 in an existing workbook. | Show how to reorder multiple worksheets in a loop using Worksheet.MoveTo in Aspose.Cells. | Explain how to verify the new worksheet order after calling MoveTo with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add worksheets, and reposition a sheet using the Worksheet.MoveTo method (C#). The example moves "MovedSheet" to index 1 and saves the file as MovedWorksheetDemo.xlsx.
    public class MoveWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add initial worksheets
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Add a worksheet that will be moved
                Worksheet movedSheet = workbook.Worksheets.Add("MovedSheet");

                // Move the worksheet to the desired index (e.g., position 1)
                movedSheet.MoveTo(1);

                // Save the workbook
                workbook.Save("MovedWorksheetDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MoveWorksheetDemo.Run();
        }
    }
}
