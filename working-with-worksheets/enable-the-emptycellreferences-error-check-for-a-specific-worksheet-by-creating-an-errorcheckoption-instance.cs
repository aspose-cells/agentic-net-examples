using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class EnableEmptyCellReferenceErrorCheck
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the collection of error‑check options for the worksheet
            ErrorCheckOptionCollection errorCheckOptions = worksheet.ErrorCheckOptions;

            // Add a new ErrorCheckOption to the collection
            int optionIndex = errorCheckOptions.Add();
            ErrorCheckOption errorCheckOption = errorCheckOptions[optionIndex];

            // Enable the EmptyCellRef error check (shows green triangle when a formula refers to an empty cell)
            errorCheckOption.SetErrorCheck(ErrorCheckType.EmptyCellRef, true);

            // Apply the error‑check option to the entire used range of the worksheet
            int maxRow = worksheet.Cells.MaxRow;
            int maxCol = worksheet.Cells.MaxDataColumn;
            CellArea fullRange = CellArea.CreateCellArea(0, 0, maxRow, maxCol);
            errorCheckOption.AddRange(fullRange);

            // Save the workbook
            string outputPath = "EnableEmptyCellReferenceErrorCheck.xlsx";
            workbook.Save(outputPath);
        }
    }
}