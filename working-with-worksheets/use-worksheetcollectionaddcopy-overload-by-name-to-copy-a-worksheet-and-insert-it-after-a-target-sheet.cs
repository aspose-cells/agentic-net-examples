using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CopyWorksheetAfterTargetDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add the source worksheet that will be copied
                Worksheet originalSheet = workbook.Worksheets.Add("Original");
                originalSheet.Cells["A1"].PutValue("Data in the original sheet");

                // Add the target worksheet after which the copy will be placed
                Worksheet targetSheet = workbook.Worksheets.Add("Target");
                targetSheet.Cells["A1"].PutValue("Data in the target sheet");

                // Copy the worksheet named "Original"
                int copiedIndex = workbook.Worksheets.AddCopy("Original");
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

                // Move the copied worksheet to be right after the target worksheet
                copiedSheet.MoveTo(targetSheet.Index + 1);

                // Optionally rename the copied sheet
                copiedSheet.Name = "OriginalCopy";

                // Define output file path
                string outputPath = "CopyAfterTarget.xlsx";

                // Save the workbook (overwrite if it already exists)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyWorksheetAfterTargetDemo.Run();
        }
    }
}