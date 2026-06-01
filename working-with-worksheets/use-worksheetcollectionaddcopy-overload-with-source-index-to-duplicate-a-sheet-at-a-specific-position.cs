using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DuplicateSheetAtSpecificPosition
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Prepare the source worksheet
                Worksheet sourceSheet = workbook.Worksheets[0];
                sourceSheet.Name = "SourceSheet";
                sourceSheet.Cells["A1"].PutValue("Data in source sheet");

                // Add another worksheet to have a place for the copied sheet
                Worksheet otherSheet = workbook.Worksheets.Add("OtherSheet");
                otherSheet.Cells["A1"].PutValue("Data in other sheet");

                // Duplicate the source sheet using AddCopy overload with source index
                int copiedIndex = workbook.Worksheets.AddCopy(0); // copies sheet at index 0
                Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
                copiedSheet.Name = "CopiedSheet";

                // Move the copied sheet to the desired position (e.g., index 1)
                copiedSheet.MoveTo(1);

                // Save the workbook
                string outputPath = "DuplicatedSheetAtPosition.xlsx";
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
            DuplicateSheetAtSpecificPosition.Run();
        }
    }
}