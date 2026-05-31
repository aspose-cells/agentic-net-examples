using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class ValidationFromTableColumn
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the table (column A)
                worksheet.Cells["A1"].PutValue("Item");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["A4"].PutValue("Cherry");
                worksheet.Cells["A5"].PutValue("Date");

                // Add a ListObject (Excel table) that covers the data range A1:A5
                int tableIndex = worksheet.ListObjects.Add("A1", "A5", true);
                ListObject table = worksheet.ListObjects[tableIndex];
                table.DisplayName = "FruitTable";

                // Define the cell where the validation will be applied (e.g., B1)
                CellArea validationArea = new CellArea
                {
                    StartRow = 0,      // Row 1 (zero‑based)
                    StartColumn = 1,   // Column B (zero‑based)
                    EndRow = 0,
                    EndColumn = 1
                };

                // Add a validation to the worksheet for the specified cell area
                int validationIdx = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIdx];

                // Configure the validation as a List type using a structured reference to the table column
                validation.Type = ValidationType.List;
                validation.Formula1 = $"{table.DisplayName}[Item]";
                validation.InCellDropDown = true;

                // Save the workbook
                string outputPath = "ValidationFromTableColumn.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ValidationFromTableColumn.Run();
        }
    }
}