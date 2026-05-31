using System;
using Aspose.Cells;

namespace AsposeCellsAutoFitDemo
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Configure load options with OnlyAuto enabled
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.AutoFitterOptions = new AutoFitterOptions
            {
                OnlyAuto = true
            };

            // Load the workbook using the configured options
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Ensure rows are auto‑fitted (optional, as OnlyAuto was applied on load)
                sheet.AutoFitRows();

                // Determine the last row that contains data
                int lastDataRow = sheet.Cells.MaxDataRow;

                // Verify each row's height and whether it matches the default font height
                for (int rowIndex = 0; rowIndex <= lastDataRow; rowIndex++)
                {
                    Row row = sheet.Cells.Rows[rowIndex];
                    Console.WriteLine($"Sheet: {sheet.Name}, Row: {rowIndex}, Height: {row.Height}, IsHeightMatched: {row.IsHeightMatched}");
                }
            }

            // Save the workbook to verify that the auto‑fit settings persist
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}