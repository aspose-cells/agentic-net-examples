using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsNamedRangeFromValidations
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Add sample data validations to demonstrate the logic
            // -------------------------------------------------
            ValidationCollection validations = sheet.Validations;

            // Validation 1: Whole number between 1 and 10 on A1:A5
            Validation v1 = validations[validations.Add(CellArea.CreateCellArea(0, 0, 4, 0))];
            v1.Type = ValidationType.WholeNumber;
            v1.Operator = OperatorType.Between;
            v1.Formula1 = "1";
            v1.Formula2 = "10";

            // Validation 2: List on C3:C4
            Validation v2 = validations[validations.Add(CellArea.CreateCellArea(2, 2, 3, 2))];
            v2.Type = ValidationType.List;
            v2.Formula1 = "Red,Green,Blue";

            // Validation 3: Custom formula on E1
            Validation v3 = validations[validations.Add(CellArea.CreateCellArea(0, 4, 0, 4))];
            v3.Type = ValidationType.Custom;
            v3.Formula1 = "=LEN(A1)>3";

            // -------------------------------------------------
            // Collect all validation areas and build a RefersTo formula
            // -------------------------------------------------
            List<string> areaRefs = new List<string>();

            foreach (Validation val in validations)
            {
                foreach (CellArea area in val.Areas)
                {
                    // Convert start and end cells to absolute A1 style (e.g., $A$1)
                    string start = "$" + CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                    string end = "$" + CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                    string areaRef = $"{sheet.Name}!{start}:{end}";
                    areaRefs.Add(areaRef);
                }
            }

            // Join multiple areas with commas – this creates a multi‑area named range
            string refersTo = "=" + string.Join(",", areaRefs);

            // -------------------------------------------------
            // Create the named range that includes only the validation cells
            // -------------------------------------------------
            int nameIndex = workbook.Worksheets.Names.Add("ValidatedCells");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            namedRange.RefersTo = refersTo;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ValidatedCellsNamedRange.xlsx");
        }
    }
}