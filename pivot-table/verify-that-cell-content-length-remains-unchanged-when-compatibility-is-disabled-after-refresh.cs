// Title: Check Pivot Table Cell Length Remains Same After Disabling Excel 2003 Compatibility (Aspose.Cells for .NET)
// Description: Creates a workbook with long text (>255 characters), builds a pivot table, records the string length of a pivot data cell after the first refresh (Excel 2003 compatibility enabled), disables IsExcel2003Compatible, refreshes again, and verifies the length is unchanged. The example also saves and reloads the file to show the compatibility flag persists.
// Keywords: Aspose.Cells | C# | pivot table | IsExcel2003Compatible | cell length verification | string truncation | Excel 2003 compatibility | refresh pivot | save workbook | load workbook
// Common Searches: Aspose.Cells pivot table length unchanged after disabling compatibility | IsExcel2003Compatible effect on long text in pivot tables | verify no truncation when turning off Excel 2003 compatibility in Aspose.Cells | persist pivot compatibility setting after saving workbook | C# example for checking pivot cell string length in Aspose.Cells
// Developer Intent: Confirm that a pivot table cell containing long text retains its original length when IsExcel2003Compatible is set to false after an initial refresh.
// Use Cases: Automated testing to ensure long strings (>255 chars) are not truncated after changing compatibility mode. | Demonstrating that the IsExcel2003Compatible property is stored with the workbook and restored on load. | Comparing cell content before and after toggling compatibility to detect unintended data loss.
// AI Prompts: Generate C# code using Aspose.Cells that creates a pivot table, toggles IsExcel2003Compatible, refreshes the pivot, and asserts the data cell length stays the same. | Write an NUnit test that validates a pivot table's string length remains unchanged after disabling Excel 2003 compatibility in Aspose.Cells. | Explain how Aspose.Cells handles long text in pivot tables when IsExcel2003Compatible is true versus false, including any truncation behavior.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with long text (>255 characters), builds a pivot table, records the string length of a pivot data cell after the first refresh (Excel 2003 compatibility enabled), disables IsExcel2003Compatible, refreshes again, and verifies the length is unchanged. The example also saves and reloads the file to show the compatibility flag persists.
class VerifyPivotCompatibility
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add a data worksheet
            Workbook wb = new Workbook();
            Worksheet dataSheet = wb.Worksheets[0];
            dataSheet.Name = "Data";

            // Add headers
            dataSheet.Cells["A1"].Value = "Category";
            dataSheet.Cells["B1"].Value = "Description";

            // Create a long description (>255 characters)
            string longDesc = new string('x', 300);

            // Add sample rows
            dataSheet.Cells["A2"].Value = "Item1";
            dataSheet.Cells["B2"].Value = longDesc;
            dataSheet.Cells["A3"].Value = "Item2";
            dataSheet.Cells["B3"].Value = longDesc;

            // Add a pivot sheet
            Worksheet pivotSheet = wb.Worksheets.Add("Pivot");

            // Create a pivot table based on the data range (include sheet name in source range)
            int pivotIndex = pivotSheet.PivotTables.Add("PivotTable", "Data!A1:B3", "A5", true);
            PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields: Row = Category, Data = Description (as count)
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Description column

            // First refresh with default compatibility (IsExcel2003Compatible = true)
            pivot.RefreshData();
            pivot.CalculateData();

            // Retrieve the length of a data cell after the first refresh
            // Pivot data starts at cell B6 in this simple layout
            Cell dataCell = pivotSheet.Cells["B6"];
            int lengthBefore = dataCell.StringValue.Length;
            Console.WriteLine("Length after first refresh (compatibility true): " + lengthBefore);

            // Disable Excel 2003 compatibility
            pivot.IsExcel2003Compatible = false;

            // Refresh again after disabling compatibility
            pivot.RefreshData();
            pivot.CalculateData();

            // Retrieve the length again; it should remain unchanged
            int lengthAfter = dataCell.StringValue.Length;
            Console.WriteLine("Length after second refresh (compatibility false): " + lengthAfter);
            Console.WriteLine("Length unchanged: " + (lengthBefore == lengthAfter));

            // Save the workbook (create rule)
            string filePath = "PivotCompatibilityDemo.xlsx";
            wb.Save(filePath);

            // Load the workbook (load rule) to demonstrate that the setting persisted
            if (File.Exists(filePath))
            {
                LoadOptions loadOptions = new LoadOptions();
                Workbook loadedWb = new Workbook(filePath, loadOptions);
                PivotTable loadedPivot = loadedWb.Worksheets["Pivot"].PivotTables[0];
                Console.WriteLine("Loaded pivot IsExcel2003Compatible: " + loadedPivot.IsExcel2003Compatible);
            }
            else
            {
                Console.WriteLine("Error: File not found after saving: " + filePath);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
