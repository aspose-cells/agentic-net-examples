using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsWorksheetUnhideDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing workbook (replace with your actual file)
            string inputPath = "HiddenSheetWorkbook.xlsx";

            try
            {
                // Verify that the workbook file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook from disk
                Workbook workbook = new Workbook(inputPath);

                // Ensure the worksheet index exists
                if (workbook.Worksheets.Count <= 1)
                {
                    Console.WriteLine("The workbook does not contain a second worksheet to unhide.");
                    return;
                }

                // Index or name of the worksheet that is hidden (second worksheet, index 1)
                Worksheet hiddenSheet = workbook.Worksheets[1];

                // Unhide the worksheet
                hiddenSheet.IsVisible = true; // equivalent to VisibilityType.Visible

                // Change the tab color of the now-visible worksheet
                hiddenSheet.TabColor = Color.Orange; // any System.Drawing.Color

                // Save the modified workbook back to disk (overwrites the original file)
                workbook.Save(inputPath);

                Console.WriteLine($"Worksheet '{hiddenSheet.Name}' is now visible with tab color set to {hiddenSheet.TabColor}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}