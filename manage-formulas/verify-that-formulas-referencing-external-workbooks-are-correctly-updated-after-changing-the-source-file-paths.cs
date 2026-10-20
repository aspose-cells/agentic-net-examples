// Title: How to verify that formulas referencing an external workbook are updated after renaming the source file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# update external workbook link after moving source file | Recalculate Excel formula that points to a renamed workbook using Aspose.Cells | Programmatically replace external reference path in Aspose.Cells formula | Validate result of external reference after source workbook rename in .NET | Change external workbook reference in saved Excel file with Aspose.Cells API
// Tags: external reference path update Aspose.Cells | replace workbook name in Excel formula C# | recalculate formulas after external link change Aspose.Cells | validate external link result Aspose.Cells .NET | manage external workbook references Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a source workbook (source1.xlsx) containing a value, builds a target workbook with a formula that references the source, moves the source file to source2.xlsx, reloads the target workbook, replaces the old file name in the formula, recalculates the workbook, and confirms that the cell value reflects the updated external reference (expected value: 10).
class Program
{
    static void Main()
    {
        try
        {
            // -------------------------------------------------
            // 1. Create the source workbook (source1.xlsx)
            // -------------------------------------------------
            var sourceWb = new Workbook();
            var sourceWs = sourceWb.Worksheets[0];
            sourceWs.Cells["A1"].PutValue(10); // A1 = 10

            const string sourceFile1 = "source1.xlsx";
            sourceWb.Save(sourceFile1);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error creating source workbook: " + ex.Message);
            return;
        }

        try
        {
            // -------------------------------------------------
            // 2. Create the target workbook that references the source workbook
            // -------------------------------------------------
            var targetWb = new Workbook();
            var targetWs = targetWb.Worksheets[0];
            // Formula that points to source1.xlsx
            targetWs.Cells["A1"].Formula = "='[source1.xlsx]Sheet1'!A1";

            const string targetFile = "target.xlsx";
            targetWb.Save(targetFile);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error creating target workbook: " + ex.Message);
            return;
        }

        try
        {
            // -------------------------------------------------
            // 3. Simulate moving/renaming the source workbook
            // -------------------------------------------------
            const string sourceFile1 = "source1.xlsx";
            const string sourceFile2 = "source2.xlsx";

            if (File.Exists(sourceFile1))
            {
                // Ensure the destination does not already exist
                if (File.Exists(sourceFile2))
                {
                    File.Delete(sourceFile2);
                }

                File.Move(sourceFile1, sourceFile2);
            }
            else
            {
                Console.WriteLine($"File '{sourceFile1}' not found. Skipping move operation.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error moving source workbook: " + ex.Message);
            return;
        }

        try
        {
            // -------------------------------------------------
            // 4. Load the target workbook again and update the external reference
            // -------------------------------------------------
            const string targetFile = "target.xlsx";
            if (!File.Exists(targetFile))
            {
                Console.WriteLine($"Target file '{targetFile}' not found.");
                return;
            }

            var loadedTarget = new Workbook(targetFile);

            // Update the external reference manually (ChangeExternalReference not available in this version)
            var ws = loadedTarget.Worksheets[0];
            string oldFormula = ws.Cells["A1"].Formula;
            if (!string.IsNullOrEmpty(oldFormula))
            {
                ws.Cells["A1"].Formula = oldFormula.Replace("source1.xlsx", "source2.xlsx");
            }

            // Recalculate formulas so the new external value is taken into account
            loadedTarget.CalculateFormula();

            // -------------------------------------------------
            // 5. Verify that the formula result reflects the updated source
            // -------------------------------------------------
            object result = ws.Cells["A1"].Value;
            Console.WriteLine("Result after updating external reference: " + result);
            // Expected output: 10
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error processing target workbook: " + ex.Message);
        }
    }
}
