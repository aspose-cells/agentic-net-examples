using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CopyWorksheetInsertAfterOriginal
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first (original) worksheet and add some data
                Worksheet originalSheet = workbook.Worksheets[0];
                originalSheet.Name = "Original";
                originalSheet.Cells["A1"].PutValue("Data in the original sheet");

                // Copy the original worksheet; AddCopy returns the index of the new copy
                int copiedIndex = workbook.Worksheets.AddCopy(0);

                // Move the copied worksheet so that it is placed directly after the original sheet
                workbook.Worksheets[copiedIndex].MoveTo(originalSheet.Index + 1);

                // Optionally rename the copied sheet
                workbook.Worksheets[originalSheet.Index + 1].Name = "CopyAfterOriginal";

                // Save the workbook
                string outputPath = "CopyAfterOriginal.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}