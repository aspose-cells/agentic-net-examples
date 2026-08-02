// Title: Create an in‑cell drop‑down list in O1 from a lookup table using Aspose.Cells for .NET (C#)
// Description: The example builds a new workbook, writes fruit names to A2:A6, adds a List‑type data validation to cell O1 that references $A$2:$A$6, enables the in‑cell drop‑down, places a label in N1, and saves the file as DropDownExample.xlsx.
// Keywords: Aspose.Cells | Aspose.Cells for .NET | C# data validation | Excel drop‑down list | lookup table | in‑cell dropdown | validation list example | cell O1 dropdown | range A2:A6 | Aspose.Cells tutorial
// Common Searches: how to add a drop‑down list to a cell with Aspose.Cells C# | Aspose.Cells data validation list referencing a range | create in‑cell dropdown in O1 using Aspose.Cells for .NET | C# example for Excel drop‑down from lookup table Aspose | Aspose.Cells add validation list programmatically
// Developer Intent: Add a List‑type data validation to cell O1 that pulls its items from a lookup range in column A.
// Use Cases: Provide users with a predefined list of options (e.g., fruit selection) for consistent data entry. | Centralize selectable values in a single range so updates automatically reflect in all linked drop‑downs. | Standardize input across multiple worksheets by reusing the same validation list.
// AI Prompts: Generate C# code that applies the same drop‑down validation to cells O1, P1, and Q1 using the lookup range A2:A6 in Aspose.Cells. | Show how to replace the fixed range $A$2:$A$6 with a named range in the validation formula for Aspose.Cells. | Explain how to programmatically modify the lookup table values and refresh the dropdown without recreating the workbook. | Create an Aspose.Cells example that reads the lookup list from an external CSV file and uses it for a cell's drop‑down.

using System;
using Aspose.Cells;

namespace AsposeCellsDropDownExample
{
    // The example builds a new workbook, writes fruit names to A2:A6, adds a List‑type data validation to cell O1 that references $A$2:$A$6, enables the in‑cell drop‑down, places a label in N1, and saves the file as DropDownExample.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Prepare a lookup table with sample values.
                //    The table will be placed in column A, rows 2-6.
                // -------------------------------------------------
                string[] lookupValues = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
                for (int i = 0; i < lookupValues.Length; i++)
                {
                    // Cells are zero‑based: row i+1 (since row 0 is header), column 0 (A)
                    sheet.Cells[i + 1, 0].PutValue(lookupValues[i]);
                }

                // -------------------------------------------------
                // 2. Create data validation for cell O1 (column 14, row 0)
                //    and link it to the lookup range A2:A6.
                // -------------------------------------------------
                // Define the target cell area (O1)
                CellArea targetArea = new CellArea
                {
                    StartRow = 0,
                    EndRow = 0,
                    StartColumn = 14,
                    EndColumn = 14
                };

                // Add a new validation rule for the target area
                int validationIndex = sheet.Validations.Add(targetArea);
                Validation validation = sheet.Validations[validationIndex];

                // Set validation properties
                validation.Type = ValidationType.List;
                validation.InCellDropDown = true;
                validation.Formula1 = "$A$2:$A$6";

                // -------------------------------------------------
                // 3. (Optional) Add a label to indicate the purpose of O1
                // -------------------------------------------------
                sheet.Cells["N1"].PutValue("Select Fruit:");

                // -------------------------------------------------
                // 4. Save the workbook to a file
                // -------------------------------------------------
                workbook.Save("DropDownExample.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
