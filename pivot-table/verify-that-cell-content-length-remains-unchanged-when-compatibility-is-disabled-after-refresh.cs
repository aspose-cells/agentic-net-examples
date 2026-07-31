// Title: Aspose.Cells C# – Verify No Truncation of >255‑Char Text After Disabling Excel 2003 Compatibility in a PivotTable
// Description: This example creates a workbook with a 300‑character string, builds a PivotTable from the data, sets IsExcel2003Compatible = false before refreshing the pivot, saves and reloads the file, then compares the original and loaded string lengths to prove that long text is not truncated.
// Keywords: Aspose.Cells | C# | PivotTable | IsExcel2003Compatible | long text | text truncation | cell length verification | Excel 2003 compatibility | 300 characters | save and reload workbook | pivot refresh | regression test
// Common Searches: Aspose.Cells keep long strings in PivotTable | disable Excel2003 compatibility C# Aspose.Cells | verify cell length after pivot refresh | prevent 255‑character limit Aspose.Cells | how to test text truncation in Aspose.Cells pivot
// Developer Intent: Confirm that turning off Excel 2003 compatibility for a PivotTable does not truncate cells containing strings longer than 255 characters after the workbook is refreshed and reloaded.
// Use Cases: Automated test to ensure long description fields remain intact in generated reports. | Data migration scenario where legacy workbooks are converted and compatibility mode is removed. | Building dashboards that require full text in pivot data fields without the 255‑char limit. | Quality‑assurance validation of Aspose.Cells pivot handling after setting IsExcel2003Compatible = false.
// AI Prompts: Create a C# unit test with Aspose.Cells that asserts a 300‑character cell value is unchanged after setting IsExcel2003Compatible to false and refreshing the PivotTable. | Generate a code snippet that logs the original and loaded string lengths and raises an exception if they differ. | Explain why the IsExcel2003Compatible property must be set before calling RefreshData on a PivotTable to avoid truncation. | Provide a PowerShell script that runs the example and outputs pass/fail based on length comparison.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This example creates a workbook with a 300‑character string, builds a PivotTable from the data, sets IsExcel2003Compatible = false before refreshing the pivot, saves and reloads the file, then compares the original and loaded string lengths to prove that long text is not truncated.
    public class VerifyCellContentLengthAfterCompatibilityDisable
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // STEP 1: Prepare source data with a long text (>255 chars)
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].Value = "ID";
            dataSheet.Cells["B1"].Value = "Description";

            // Long description (300 characters)
            string longDescription = new string('x', 300);
            dataSheet.Cells["A2"].Value = 1;
            dataSheet.Cells["B2"].Value = longDescription;

            // -------------------------------------------------
            // STEP 2: Add a PivotTable based on the source data
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("PivotTable", "A1:B2", "A4", false);
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Add fields: ID as Row, Description as Data (show as Max to keep text)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // ID
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);     // Description

            // -------------------------------------------------
            // STEP 3: Disable Excel 2003 compatibility BEFORE refresh
            // -------------------------------------------------
            pivotTable.IsExcel2003Compatible = false; // Allow strings longer than 255 chars

            // Refresh the PivotTable so that any compatibility logic is applied
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // STEP 4: Save the workbook to a temporary file
            // -------------------------------------------------
            string tempFile = Guid.NewGuid().ToString() + ".xlsx";
            workbook.Save(tempFile);

            // -------------------------------------------------
            // STEP 5: Load the workbook back (ensure file exists)
            // -------------------------------------------------
            if (!File.Exists(tempFile))
                throw new FileNotFoundException("The temporary workbook file was not found.", tempFile);

            LoadOptions loadOptions = new LoadOptions();
            Workbook loadedWorkbook = new Workbook(tempFile, loadOptions);

            // -------------------------------------------------
            // STEP 6: Verify that the original cell content length is unchanged
            // -------------------------------------------------
            Worksheet loadedDataSheet = loadedWorkbook.Worksheets["Data"];
            string loadedValue = loadedDataSheet.Cells["B2"].StringValue;

            Console.WriteLine("Original length : " + longDescription.Length);
            Console.WriteLine("Loaded length   : " + loadedValue.Length);
            Console.WriteLine("Lengths equal?  : " + (longDescription.Length == loadedValue.Length));
        }
    }
}
