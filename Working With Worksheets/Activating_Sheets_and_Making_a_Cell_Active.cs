using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ActivateSheetAndCellDemo
    {
        public static void Run()
        {
            // Create a new workbook with default first worksheet
            Workbook workbook = new Workbook();

            // Add two more worksheets for demonstration
            workbook.Worksheets.Add("SecondSheet");
            workbook.Worksheets.Add("ThirdSheet");

            // Set the second worksheet as the active sheet (index is zero‑based)
            workbook.Worksheets.ActiveSheetIndex = 1;

            // Access the active worksheet
            Worksheet activeWorksheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

            // Set the active cell within the active worksheet to B2
            activeWorksheet.ActiveCell = "B2";

            // Optionally put a value in the active cell to see it after opening
            activeWorksheet.Cells["B2"].PutValue("Active");

            // Save the workbook to a temporary file
            string outputPath = Path.Combine(Path.GetTempPath(), "ActiveSheetAndCellDemo.xlsx");
            workbook.Save(outputPath);

            // Load the saved workbook to verify the active sheet and cell
            Workbook loadedWorkbook = new Workbook(outputPath);

            // Retrieve the active sheet index and active cell address
            int activeSheetIdx = loadedWorkbook.Worksheets.ActiveSheetIndex;
            string activeCellAddress = loadedWorkbook.Worksheets[activeSheetIdx].ActiveCell;

            // Output verification information
            Console.WriteLine($"Active Sheet Index: {activeSheetIdx}");
            Console.WriteLine($"Active Sheet Name: {loadedWorkbook.Worksheets[activeSheetIdx].Name}");
            Console.WriteLine($"Active Cell in Active Sheet: {activeCellAddress}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ActivateSheetAndCellDemo.Run();
        }
    }
}