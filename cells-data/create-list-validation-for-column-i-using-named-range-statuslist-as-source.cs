using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ColumnIListValidation
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate source list in column A (A1:A4)
                worksheet.Cells["A1"].PutValue("Open");
                worksheet.Cells["A2"].PutValue("In Progress");
                worksheet.Cells["A3"].PutValue("Closed");
                worksheet.Cells["A4"].PutValue("On Hold");

                // Define validation area: column I (index 8), rows 0‑1000
                int startRow = 0;
                int endRow = 1000;
                int columnI = 8; // column I (A=0, B=1, ..., I=8)

                CellArea validationArea = new CellArea
                {
                    StartRow = startRow,
                    EndRow = endRow,
                    StartColumn = columnI,
                    EndColumn = columnI
                };

                // Add list validation using a direct range reference
                int validationIndex = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIndex];
                validation.Type = ValidationType.List;
                validation.Formula1 = "=$A$1:$A$4"; // direct range for dropdown list
                validation.InCellDropDown = true;   // show dropdown arrow

                // Save the workbook
                string outputPath = "ColumnI_ListValidation.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ColumnIListValidation.Run();
        }
    }
}