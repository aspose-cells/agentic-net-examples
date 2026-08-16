// Title: Aspose.Cells .NET – Create a Named Range for Cells with Data Validation
// Description: Demonstrates how to generate a workbook, add data‑validation rules, gather all validation CellArea addresses, build a union RefersTo string, and add a global named range (ValidatedCells) that references only the validated cells before saving the file.
// Keywords: Aspose.Cells named range validation | C# collect data validation cells | union address RefersTo Aspose | global named range validated cells | extract validation areas .NET
// Common Searches: Aspose.Cells create named range for validated cells C# | how to list all cells with data validation in Aspose.Cells | build union RefersTo string from validation areas | programmatically add named range for data validation | Aspose.Cells get validation cell addresses
// Developer Intent: Programmatically define a named range that points exclusively to cells containing data‑validation rules.
// Use Cases: Reference every validated input in formulas via a single named range. | Generate an audit report of all cells that enforce data validation. | Apply protection, styling, or conditional formatting to all validated cells at once.
// AI Prompts: Write C# code using Aspose.Cells to create a named range that includes all cells with data validation, handling multiple validation areas and union references. | Explain step‑by‑step how to collect validation CellArea objects and construct the RefersTo string for a named range in Aspose.Cells. | Show how to modify the example to create a worksheet‑scoped named range instead of a global one.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsValidationNamedRange
{
    // Demonstrates how to generate a workbook, add data‑validation rules, gather all validation CellArea addresses, build a union RefersTo string, and add a global named range (ValidatedCells) that references only the validated cells before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // -------------------------------------------------
            // Sample: add some data validations to demonstrate
            // -------------------------------------------------
            ValidationCollection validations = sheet.Validations;

            // Validation 1: whole number between 1 and 10 on A1:A5
            Validation v1 = validations[validations.Add(CellArea.CreateCellArea(0, 0, 4, 0))];
            v1.Type = ValidationType.WholeNumber;
            v1.Operator = OperatorType.Between;
            v1.Formula1 = "1";
            v1.Formula2 = "10";

            // Validation 2: list on C3
            Validation v2 = validations[validations.Add(CellArea.CreateCellArea(2, 2, 2, 2))];
            v2.Type = ValidationType.List;
            v2.Formula1 = "Red,Green,Blue";

            // -------------------------------------------------
            // Collect all validation areas and build a union address string
            // -------------------------------------------------
            List<string> areaAddresses = new List<string>();

            foreach (Validation val in validations)
            {
                foreach (CellArea area in val.Areas)
                {
                    // Convert start and end cells to A1 style addresses
                    string startAddr = sheet.Cells[area.StartRow, area.StartColumn].Name;
                    string endAddr = sheet.Cells[area.EndRow, area.EndColumn].Name;

                    // Build full address with sheet name
                    string fullAddr = $"'{sheet.Name}'!{startAddr}";
                    if (startAddr != endAddr)
                        fullAddr += $":{endAddr}";

                    areaAddresses.Add(fullAddr);
                }
            }

            // If there are validation areas, create a named range that refers to their union
            if (areaAddresses.Count > 0)
            {
                // Join addresses with commas to create a union reference
                string refersTo = "=" + string.Join(",", areaAddresses);

                // Add the named range to the workbook (global scope)
                int nameIndex = workbook.Worksheets.Names.Add("ValidatedCells");
                Name namedRange = workbook.Worksheets.Names[nameIndex];
                namedRange.RefersTo = refersTo;
            }

            // Save the workbook
            workbook.Save("ValidatedCellsNamedRange.xlsx");
        }
    }
}
