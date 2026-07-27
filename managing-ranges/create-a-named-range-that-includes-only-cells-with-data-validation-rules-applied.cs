using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsValidationNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ----- Sample data validations (for demonstration) -----
            // Validation 1 on A1:A3
            Validation val1 = sheet.Validations[sheet.Validations.Add()];
            val1.Type = ValidationType.List;
            val1.Formula1 = "Option1,Option2,Option3";
            val1.AddArea(CellArea.CreateCellArea(0, 0, 2, 0)); // A1:A3

            // Validation 2 on C5
            Validation val2 = sheet.Validations[sheet.Validations.Add()];
            val2.Type = ValidationType.WholeNumber;
            val2.Operator = OperatorType.Between;
            val2.Formula1 = "1";
            val2.Formula2 = "10";
            val2.AddArea(CellArea.CreateCellArea(4, 2, 4, 2)); // C5

            // Validation 3 on E2:E4 (multiple separate areas)
            Validation val3 = sheet.Validations[sheet.Validations.Add()];
            val3.Type = ValidationType.Custom;
            val3.Formula1 = "=ISNUMBER(A1)";
            val3.AddArea(CellArea.CreateCellArea(1, 4, 1, 4)); // E2
            val3.AddArea(CellArea.CreateCellArea(2, 4, 2, 4)); // E3
            val3.AddArea(CellArea.CreateCellArea(3, 4, 3, 4)); // E4

            // ----- Collect all validation areas -----
            List<string> areaRefs = new List<string>();
            foreach (Validation validation in sheet.Validations)
            {
                foreach (CellArea area in validation.Areas)
                {
                    // Convert start and end indices to A1 style addresses
                    string startAddr = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string endAddr = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                    string fullRef = $"'{sheet.Name}'!{startAddr}:{endAddr}";
                    areaRefs.Add(fullRef);
                }
            }

            // If there are validation areas, create a named range that references only those cells
            if (areaRefs.Count > 0)
            {
                // Join the individual area references with commas to form a union range
                string refersToFormula = "=" + string.Join(",", areaRefs);

                // Add the named range to the workbook
                int nameIndex = workbook.Worksheets.Names.Add("ValidatedCells");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                namedRange.RefersTo = refersToFormula;
            }

            // Save the workbook
            workbook.Save("ValidatedCellsNamedRange.xlsx");
        }
    }
}