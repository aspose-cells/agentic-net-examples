// Title: Move a Worksheet to the End of an Aspose.Cells Workbook (C#)
// Description: Demonstrates how to create a workbook, add sheets, select a specific worksheet, compute the last zero‑based index, use Worksheet.MoveTo to place the sheet at the final position, and save the file as WorksheetMovedToEnd.xlsx.
// Keywords: Aspose.Cells move worksheet | C# Worksheet.MoveTo | reorder sheets Aspose.Cells | place sheet at end of workbook | Aspose.Cells .NET worksheet ordering | move sheet to last index
// Common Searches: Aspose.Cells move worksheet to last position C# | Worksheet.MoveTo example Aspose.Cells .NET | how to reorder worksheets in Aspose.Cells | C# code to shift a sheet to the end of a workbook | Aspose.Cells place summary sheet at end
// Developer Intent: Programmatically move a chosen worksheet to the final index of a workbook.
// Use Cases: Add a summary or index sheet after generating a report and ensure it appears last. | Insert a dynamically created analysis worksheet and position it at the end for user convenience. | Reorder template sheets so that a user‑selected sheet is always the final tab.
// AI Prompts: Write C# code using Aspose.Cells to move the worksheet named "Data" to the end of the workbook and save it as "Result.xlsx". | Explain the Worksheet.MoveTo method in Aspose.Cells, including how zero‑based indices work, with a sample moving the second sheet to the last position. | Provide a C# snippet that scans all worksheets and moves any sheet whose name starts with "Temp" to the end of the workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, add sheets, select a specific worksheet, compute the last zero‑based index, use Worksheet.MoveTo to place the sheet at the final position, and save the file as WorksheetMovedToEnd.xlsx.
    public class MoveWorksheetToEndDemo
    {
        // Entry point for the console application
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully as 'WorksheetMovedToEnd.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Add sample worksheets
            workbook.Worksheets.Add("Sheet1");
            workbook.Worksheets.Add("Sheet2");
            workbook.Worksheets.Add("Sheet3");

            // Choose the worksheet to move (e.g., "Sheet2")
            Worksheet sheetToMove = workbook.Worksheets["Sheet2"];

            // Calculate the last index (zero‑based)
            int lastIndex = workbook.Worksheets.Count - 1;

            // Move the selected worksheet to the last position
            sheetToMove.MoveTo(lastIndex);

            // Save the workbook to the current directory
            workbook.Save("WorksheetMovedToEnd.xlsx");
        }
    }
}
