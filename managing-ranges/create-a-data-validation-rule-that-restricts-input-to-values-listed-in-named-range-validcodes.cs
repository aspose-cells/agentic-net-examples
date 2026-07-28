// Title: C# – Add List Data Validation from a Named Range with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, fill cells A1:A5 with codes, define a named range called ValidCodes, apply a list‑type data validation to cell B1 that references the named range, enable the in‑cell dropdown, and save the file as ValidationWithNamedRange.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | data validation list | named range | Excel dropdown | list validation | cell validation | Aspose.Cells example
// Common Searches: Aspose.Cells list validation named range C# | how to reference a named range in data validation with Aspose.Cells | create dropdown list from named range using Aspose.Cells .NET | C# Aspose.Cells add data validation list
// Developer Intent: Apply a list‑type validation to a cell that limits entries to the values defined in a workbook named range.
// Use Cases: Ensure product codes entered in column B match a predefined list in column A. | Provide a selectable list of status options stored on a hidden sheet. | Offer a dropdown of department identifiers for data‑entry forms without hard‑coding values.
// AI Prompts: Generate C# code with Aspose.Cells that creates a named range "ValidItems" and adds a list validation with a dropdown to cell D5. | Explain how to set the Formula1 property to reference a named range in Aspose.Cells data validation and enable the in‑cell dropdown.

using System;
using Aspose.Cells;

namespace AsposeCellsValidationExample
{
    // Demonstrates how to create a workbook, fill cells A1:A5 with codes, define a named range called ValidCodes, apply a list‑type data validation to cell B1 that references the named range, enable the in‑cell dropdown, and save the file as ValidationWithNamedRange.xlsx using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // 1. Populate the range that will be used as a named range "ValidCodes"
                // -------------------------------------------------
                // Example valid codes placed in column A (A1:A5)
                worksheet.Cells["A1"].PutValue("CodeA");
                worksheet.Cells["A2"].PutValue("CodeB");
                worksheet.Cells["A3"].PutValue("CodeC");
                worksheet.Cells["A4"].PutValue("CodeD");
                worksheet.Cells["A5"].PutValue("CodeE");

                // -------------------------------------------------
                // 2. Define the named range "ValidCodes" that refers to A1:A5
                // -------------------------------------------------
                // Create a range object covering A1:A5 (rows, columns)
                Aspose.Cells.Range validCodesRange = worksheet.Cells.CreateRange(0, 0, 5, 1);

                // Add the named range to the workbook and obtain the Name object
                int nameIndex = workbook.Worksheets.Names.Add("ValidCodes");
                Name validCodesName = workbook.Worksheets.Names[nameIndex];

                // Point the name to the created range (RefersTo expects a string)
                validCodesName.RefersTo = validCodesRange.RefersTo;

                // -------------------------------------------------
                // 3. Add a data validation to cell B1 that restricts input to the named range
                // -------------------------------------------------
                // Define the area for validation (cell B1)
                CellArea validationArea = CellArea.CreateCellArea(0, 1, 0, 1); // Row 0, Column 1

                // Add the validation to the worksheet's validation collection
                int validationIndex = worksheet.Validations.Add(validationArea);
                Validation validation = worksheet.Validations[validationIndex];

                // Set validation type to List and point Formula1 to the named range
                validation.Type = ValidationType.List;
                validation.Formula1 = "=ValidCodes"; // Reference the named range
                validation.InCellDropDown = true;    // Show dropdown list in the cell

                // -------------------------------------------------
                // 4. Save the workbook
                // -------------------------------------------------
                workbook.Save("ValidationWithNamedRange.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
