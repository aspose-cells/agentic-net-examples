// Title: Copy a Table with Formulas to a New Sheet while Preserving Calculation Mode – Aspose.Cells C#
// Description: Shows how to build a source worksheet with headers, data rows and a Total formula, add a destination sheet, copy the entire sheet (including formulas) using CopyOptions.ReferToSheetWithSameName, retain the original CalcModeType, optionally recalculate formulas, and save the workbook as an Excel file.
// Keywords: Aspose.Cells C# copy worksheet | copy table with formulas | preserve calculation mode | CopyOptions ReferToSheetWithSameName | FormulaSettings CalcModeType | .NET Excel automation | duplicate sheet formulas | Aspose.Cells workbook settings | Excel formula copy example
// Common Searches: Aspose.Cells copy worksheet without changing calculation mode | How to keep formula references when duplicating a sheet in C# | CopyOptions.ReferToSheetWithSameName usage example | Preserve CalcModeType after copying a sheet Aspose.Cells | Copy table with formulas to another sheet .NET
// Developer Intent: Duplicate a worksheet that contains formulas to a new sheet while leaving the workbook's calculation mode unchanged.
// Use Cases: Generate period‑specific report tabs from a template that includes calculated columns. | Create a backup of a data‑entry sheet without triggering a full recalculation. | Clone a financial model worksheet for scenario analysis while maintaining original sheet references.
// AI Prompts: Provide C# code using Aspose.Cells to copy a worksheet with formulas to another sheet and keep the original calculation mode. | Show an example of CopyOptions.ReferToSheetWithSameName to retain formula links when copying a table in Aspose.Cells for .NET. | Write a snippet that saves the original CalcModeType, copies a sheet, recalculates formulas, and writes the workbook to disk.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyTableWithFormulas
{
    // Shows how to build a source worksheet with headers, data rows and a Total formula, add a destination sheet, copy the entire sheet (including formulas) using CopyOptions.ReferToSheetWithSameName, retain the original CalcModeType, optionally recalculate formulas, and save the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Prepare the source worksheet with a simple table
            // -------------------------------------------------
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Fill header
            sourceSheet.Cells["A1"].PutValue("Item");
            sourceSheet.Cells["B1"].PutValue("Quantity");
            sourceSheet.Cells["C1"].PutValue("Price");
            sourceSheet.Cells["D1"].PutValue("Total");

            // Fill some data rows
            sourceSheet.Cells["A2"].PutValue("Apple");
            sourceSheet.Cells["B2"].PutValue(5);
            sourceSheet.Cells["C2"].PutValue(2.0);
            sourceSheet.Cells["A3"].PutValue("Banana");
            sourceSheet.Cells["B3"].PutValue(3);
            sourceSheet.Cells["C3"].PutValue(1.5);

            // Formula column: Total = Quantity * Price
            sourceSheet.Cells["D2"].Formula = "=B2*C2";
            sourceSheet.Cells["D3"].Formula = "=B3*C3";

            // -------------------------------------------------
            // Add a new worksheet that will receive the copied table
            // -------------------------------------------------
            Worksheet destSheet = workbook.Worksheets.Add("Copy");

            // -------------------------------------------------
            // Copy the source worksheet content (including formulas) to the destination sheet
            // Use CopyOptions to keep references to sheets with the same name
            // -------------------------------------------------
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ReferToSheetWithSameName = true; // keep formula references consistent

            // The Copy method copies the whole worksheet content into the target worksheet
            destSheet.Copy(sourceSheet, copyOptions);

            // -------------------------------------------------
            // Ensure the calculation mode of the workbook remains unchanged
            // (store the original mode and reapply it after the copy operation)
            // -------------------------------------------------
            CalcModeType originalCalcMode = workbook.Settings.FormulaSettings.CalculationMode;
            workbook.Settings.FormulaSettings.CalculationMode = originalCalcMode;

            // -------------------------------------------------
            // Optionally calculate formulas so that the copied sheet shows results immediately
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("CopiedTableWithFormulas.xlsx");
        }
    }
}
