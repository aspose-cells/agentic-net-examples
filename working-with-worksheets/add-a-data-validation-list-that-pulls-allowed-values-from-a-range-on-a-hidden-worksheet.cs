// Title: Aspose.Cells .NET: Create a data‑validation drop‑down list from a hidden worksheet (C#)
// Description: This example shows how to programmatically add a hidden worksheet, fill it with option values, and attach a list‑type data validation to cell B2 on a visible sheet. The validation’s Formula1 references the hidden range (HiddenValues!$A$1:$A$5), providing an in‑cell drop‑down while keeping the source list concealed. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | data validation | list validation | hidden worksheet | drop‑down list | Excel automation | Formula1 reference | CellArea | validation range | programmatic Excel | Aspose.Cells API
// Common Searches: Aspose.Cells C# hidden sheet validation list | Create dropdown list from hidden range using Aspose.Cells | Aspose.Cells data validation list formula hidden worksheet | How to hide worksheet and use it for validation in Aspose.Cells | Aspose.Cells list validation reference another sheet
// Developer Intent: Add a list‑type data validation to a cell that pulls its allowed values from a range on a hidden worksheet.
// Use Cases: Provide users with a clean data‑entry interface while storing the source list on a hidden sheet to prevent accidental edits. | Build Excel templates where lookup tables are hidden but still drive validation rules for consistent input. | Generate workbooks programmatically that enforce predefined choices without exposing the reference data to end users.
// AI Prompts: Generate C# code with Aspose.Cells that creates a hidden worksheet, populates it with options, and applies a list validation to a cell referencing that hidden range. | Explain step‑by‑step how to hide a worksheet and reference its cells in a data‑validation formula using Aspose.Cells for .NET. | Show how to modify the validation source range dynamically based on values added to a hidden sheet in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example shows how to programmatically add a hidden worksheet, fill it with option values, and attach a list‑type data validation to cell B2 on a visible sheet. The validation’s Formula1 references the hidden range (HiddenValues!$A$1:$A$5), providing an in‑cell drop‑down while keeping the source list concealed. The workbook is saved as an XLSX file.
    public class ValidationFromHiddenSheet
    {
        public static void Main(string[] args)
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

            // -------------------------------------------------
            // 1. Prepare the hidden worksheet with allowed values
            // -------------------------------------------------
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenValues");
            hiddenSheet.IsVisible = false; // hide the sheet

            // Populate the hidden range A1:A5 with list items
            for (int i = 0; i < 5; i++)
            {
                hiddenSheet.Cells[i, 0].PutValue($"Option{i + 1}");
            }

            // -------------------------------------------------
            // 2. Set up the visible worksheet where validation will be applied
            // -------------------------------------------------
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "DataEntry";

            // Define the cell (e.g., B2) that will have the drop‑down list
            CellArea validationArea = CellArea.CreateCellArea(1, 1, 1, 1); // Row 2, Column 2 (B2)

            // Add a new validation to the collection for the defined area
            int validationIndex = visibleSheet.Validations.Add(validationArea);
            Validation validation = visibleSheet.Validations[validationIndex];

            // Configure the validation as a List that references the hidden range
            validation.Type = ValidationType.List;
            validation.Formula1 = "HiddenValues!$A$1:$A$5"; // reference to hidden sheet
            validation.InCellDropDown = true; // show the drop‑down arrow

            // -------------------------------------------------
            // 3. Save the workbook
            // -------------------------------------------------
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "DataValidationFromHiddenSheet.xlsx");
            workbook.Save(outputPath);
        }
    }
}
