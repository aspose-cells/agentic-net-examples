using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorksheetRelocateDemo
    {
        public static void Run()
        {
            // Load an existing workbook from disk
            Workbook workbook = new Workbook("input.xlsx");

            // Identify the worksheet to relocate (by name or index)
            Worksheet sheetToMove = workbook.Worksheets["Sheet3"];
            if (sheetToMove == null)
            {
                // If the sheet does not exist, create it and give it the desired name
                int newIndex = workbook.Worksheets.Add();
                sheetToMove = workbook.Worksheets[newIndex];
                sheetToMove.Name = "Sheet3";
            }

            // Move the worksheet to the desired position.
            // Index 0 places it as the first sheet in the workbook.
            sheetToMove.MoveTo(0);

            // Save the workbook with the updated sheet order
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorksheetRelocateDemo.Run();
        }
    }
}