using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportHiddenWorksheetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add data to the default visible sheet
                Workbook workbook = new Workbook();
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Data from visible sheet");

                // Add a second worksheet and hide it
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Data from hidden sheet");
                hiddenSheet.IsVisible = false; // Mark the worksheet as hidden

                // Configure HTML save options to include hidden worksheets
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true,          // Ensure hidden sheets are exported
                    ExportActiveWorksheetOnly = false      // Export the whole workbook
                };

                string outputPath = "ExportHiddenWorksheet_Enabled.html";

                // Save the workbook to HTML; hidden worksheet will be included in the output
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHiddenWorksheetDemo.Run();
        }
    }
}