// Title: Lock formula cells and protect worksheets in an Excel file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that iterates through every worksheet in a Workbook, sets the IsLocked style on cells that contain a formula, and then calls Protect on each sheet with Aspose.Cells. | Create a reusable method `LockFormulaCells(Workbook wb)` that locks all formula‑containing cells and applies full protection to each worksheet before saving. | Write a snippet that loads an existing .xlsx file, locks only the cells with formulas, protects the worksheets, and saves the result using Aspose.Cells for .NET.
// Common Searches: asp.net lock cells that have formulas with Aspose.Cells | c# protect worksheet after locking formula cells using Aspose.Cells | how to set IsLocked property for formula cells in an Excel workbook with Aspose.Cells | iterate all worksheets and lock formula cells before saving in C# Aspose.Cells
// Tags: apply IsLocked style to formula cells Aspose.Cells | worksheet protection with Aspose.Cells C# | formula detection and style update Aspose.Cells | secure Excel formulas via style locking C# | bulk cell style modification Aspose.Cells

using Aspose.Cells;

// The program loads an existing Excel workbook, walks through each worksheet and cell, applies the IsLocked flag to any cell that contains a formula, protects the worksheet so locked cells cannot be edited, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the cells collection of the current worksheet
            Cells cells = sheet.Cells;

            // Loop through each cell in the worksheet
            foreach (Cell cell in cells)
            {
                // If the cell contains a formula, lock it
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Retrieve the current style of the cell
                    Style style = cell.GetStyle();

                    // Set the IsLocked property to true
                    style.IsLocked = true;

                    // Apply the updated style back to the cell
                    cell.SetStyle(style);
                }
            }

            // Protect the worksheet so that locked cells cannot be edited
            sheet.Protect(ProtectionType.All);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
