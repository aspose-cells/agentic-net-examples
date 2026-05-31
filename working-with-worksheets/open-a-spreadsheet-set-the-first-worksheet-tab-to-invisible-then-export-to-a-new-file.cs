using System;
using Aspose.Cells;

namespace AsposeCellsVisibilityDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing spreadsheet
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Hide the first worksheet tab (index 0)
            // You can use IsVisible, SetVisible, or VisibilityType. Here we use IsVisible.
            workbook.Worksheets[0].IsVisible = false;

            // Path for the exported workbook
            string outputPath = "output.xlsx";

            // Save the workbook to a new file; hidden sheet will remain hidden in the saved file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to '{outputPath}' with the first worksheet hidden.");
        }
    }
}