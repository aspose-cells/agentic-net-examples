using System;
using Aspose.Cells;

namespace AsposeCellsExportToXls
{
    class Program
    {
        static void Main()
        {
            // Load the merged workbook (replace with actual path if needed)
            // If the workbook is already created in memory, you can skip this step.
            Workbook mergedWorkbook = new Workbook("mergedWorkbook.xlsx");

            // Create XLS save options for Excel 97‑2003 compatibility
            XlsSaveOptions xlsOptions = new XlsSaveOptions();

            // Optional: configure options as required
            xlsOptions.MatchColor = true;               // Match font colors to the 56‑color palette
            xlsOptions.ValidateMergedAreas = true;      // Validate merged cells before saving
            xlsOptions.MergeAreas = true;               // Merge conditional formatting/validation areas
            xlsOptions.SortNames = true;                // Sort defined names

            // Save the workbook as an XLS file using the save options
            mergedWorkbook.Save("mergedWorkbook.xls", xlsOptions);
        }
    }
}