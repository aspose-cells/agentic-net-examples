using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class InsertWorksheetAfterSpecifiedSheet
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook with a default worksheet
            Workbook workbook = new Workbook();

            // Add sample worksheets
            workbook.Worksheets.Add("Report");
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Summary");

            // Name of the sheet after which we want to place the new sheet
            string targetSheetName = "Data";

            // Retrieve the target worksheet
            Worksheet targetSheet = workbook.Worksheets[targetSheetName];

            // Determine the index where the new sheet should be inserted (after the target sheet)
            int insertIndex = targetSheet.Index + 1;

            // Insert a new worksheet at the calculated position
            Worksheet newSheet = workbook.Worksheets.Insert(insertIndex, SheetType.Worksheet, "InsertedAfterData");

            // Add some data to the newly inserted sheet
            newSheet.Cells["A1"].PutValue($"This sheet was inserted after '{targetSheetName}'.");

            // Save the workbook
            string outputPath = "WorksheetInsertedAfterSpecifiedSheet.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}