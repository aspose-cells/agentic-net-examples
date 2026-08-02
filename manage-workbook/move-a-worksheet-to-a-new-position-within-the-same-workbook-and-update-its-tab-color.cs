using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MoveWorksheetAndSetTabColor
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add three worksheets with custom names
                workbook.Worksheets.Add("Sheet1");
                workbook.Worksheets.Add("Sheet2");
                workbook.Worksheets.Add("Sheet3");

                // Add a new worksheet that we will move
                Worksheet movedSheet = workbook.Worksheets.Add("MovedSheet");

                // Move the worksheet to index 1 (second position)
                movedSheet.MoveTo(1);

                // Change the tab color of the moved worksheet to Green
                movedSheet.TabColor = Color.Green;

                // Define output file path
                string outputPath = "MovedSheetWithTabColor.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            MoveWorksheetAndSetTabColor.Run();
        }
    }
}