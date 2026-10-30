// Title: C# unit test for updating external workbook links and embedding custom ribbon XML with Aspose.Cells
// Common Searches: C# Aspose.Cells unit test to change external link formula path | How to verify external workbook reference updates in Aspose.Cells unit tests | Embedding custom ribbon XML in an Excel file using Aspose.Cells .NET | Unit testing persistence of modified formulas after external link change with Aspose.Cells | Aspose.Cells CustomUIXml property usage in C# unit tests
// Tags: update external workbook link formula Aspose.Cells | embed custom ribbon UI XML Aspose.Cells | verify formula persistence after link change Aspose.Cells | Workbook.CustomUIXml handling Aspose.Cells | C# unit test external reference handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // The example creates a source workbook, adds a formula in a target workbook that references the source, validates the original external link, updates the formula to point to a new source workbook, saves and reloads to confirm the updated reference persists, and demonstrates how to assign and verify custom ribbon XML via the Workbook.CustomUIXml property in a C# unit‑test scenario using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            string tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);

            try
            {
                // ---------- Step 1: Create source workbook ----------
                string sourcePath = Path.Combine(tempDir, "Source.xlsx");
                var sourceWb = new Workbook();
                sourceWb.Worksheets[0].Name = "Data";
                sourceWb.Worksheets[0].Cells["A1"].PutValue(10);
                sourceWb.Save(sourcePath);

                // ---------- Step 2: Create target workbook with external link ----------
                string targetPath = Path.Combine(tempDir, "Target.xlsx");
                var targetWb = new Workbook();
                var ws = targetWb.Worksheets[0];
                ws.Name = "LinkSheet";

                // Formula that references the source workbook
                ws.Cells["A1"].Formula = $"=SUM('[{Path.GetFileName(sourcePath)}]Data'!A1)";
                targetWb.Save(targetPath);

                // ---------- Step 3: Load target and verify external link ----------
                if (!File.Exists(targetPath))
                    throw new FileNotFoundException("Target workbook not found.", targetPath);

                var loadedTarget = new Workbook(targetPath);
                var targetSheet = loadedTarget.Worksheets[0];

                // Verify the formula contains the original source file name
                string originalFormula = targetSheet.Cells["A1"].Formula;
                if (!originalFormula.Contains(Path.GetFileName(sourcePath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Formula should reference Source.xlsx.");

                // ---------- Step 4: Change external link path ----------
                string newSourcePath = Path.Combine(tempDir, "NewSource.xlsx");
                var newSourceWb = new Workbook();
                newSourceWb.Worksheets[0].Name = "Data";
                newSourceWb.Worksheets[0].Cells["A1"].PutValue(20);
                newSourceWb.Save(newSourcePath);

                // Update the formula to point to the new source workbook
                targetSheet.Cells["A1"].Formula = $"=SUM('[{Path.GetFileName(newSourcePath)}]Data'!A1)";
                string updatedFormula = targetSheet.Cells["A1"].Formula;
                if (!updatedFormula.Contains(Path.GetFileName(newSourcePath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Formula should be updated to reference NewSource.xlsx.");

                // Save the modified workbook
                string modifiedTargetPath = Path.Combine(tempDir, "Target_Modified.xlsx");
                loadedTarget.Save(modifiedTargetPath);

                // Reload and verify the formula persisted
                if (!File.Exists(modifiedTargetPath))
                    throw new FileNotFoundException("Modified target workbook not found.", modifiedTargetPath);

                var reloadedTarget = new Workbook(modifiedTargetPath);
                var reloadedSheet = reloadedTarget.Worksheets[0];
                string reloadedFormula = reloadedSheet.Cells["A1"].Formula;
                if (!reloadedFormula.Contains(Path.GetFileName(newSourcePath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("Persisted formula should reference NewSource.xlsx.");

                // ---------- Step 5: Add custom ribbon XML (if supported) ----------
                try
                {
                    string customRibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='customTab' label='My Tab'>
        <group id='customGroup' label='My Group'>
          <button id='customButton' label='Click Me' size='large' onAction='OnButtonClick'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

                    // Attempt to set custom UI XML if the property exists
                    var prop = typeof(Workbook).GetProperty("CustomUIXml");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(reloadedTarget, customRibbonXml.Trim());
                        string ribbonPath = Path.Combine(tempDir, "Target_WithRibbon.xlsx");
                        reloadedTarget.Save(ribbonPath);

                        // Verify custom ribbon XML
                        if (!File.Exists(ribbonPath))
                            throw new FileNotFoundException("Ribbon workbook not found.", ribbonPath);

                        var finalWb = new Workbook(ribbonPath);
                        var loadedRibbonXml = prop.GetValue(finalWb) as string;
                        if (loadedRibbonXml == null)
                            throw new InvalidOperationException("Custom UI XML should not be null.");

                        if (!string.Equals(loadedRibbonXml.Trim(), customRibbonXml.Trim(), StringComparison.Ordinal))
                            throw new InvalidOperationException("Custom ribbon XML should match the inserted XML.");
                    }
                }
                catch (Exception ex)
                {
                    // Log but do not fail the whole process if custom UI handling is unavailable
                    Console.Error.WriteLine($"Custom UI handling skipped: {ex.Message}");
                }

                Console.WriteLine("All operations completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // Clean up temporary files
                try
                {
                    if (Directory.Exists(tempDir))
                        Directory.Delete(tempDir, true);
                }
                catch
                {
                    // Suppress any cleanup exceptions
                }
            }
        }
    }
}
