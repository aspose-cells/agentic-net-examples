// Title: Delete blank rows on the first worksheet while preserving formulas using DeleteOptions.UpdateReference in Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes all empty rows from the first worksheet of a workbook and updates any cell references by setting DeleteOptions.UpdateReference = true. | Show how to configure DeleteOptions in Aspose.Cells to delete blank rows and keep formulas correct in a .NET application. | Provide a step‑by‑step example that creates a workbook, inserts data with gaps, deletes blank rows with reference updates, and saves the result.
// Common Searches: asp.net delete blank rows first worksheet deleteoptions.updatereference true | c# aspocells remove empty rows and keep formulas intact | how to use DeleteBlankRows with UpdateReference option in Aspose.Cells | preserve cell references when deleting rows in Aspose.Cells .NET
// Tags: Aspose.Cells DeleteBlankRows with UpdateReference | C# delete empty rows preserving formulas | DeleteOptions.UpdateReference usage Aspose.Cells | first worksheet blank row removal .NET | Aspose.Cells row deletion reference update

using System;
using Aspose.Cells;

namespace AsposeCellsDeleteBlankRowsExample
{
    // The example creates a workbook, adds data with intermittent blank rows, configures DeleteOptions with UpdateReference=true, deletes all blank rows on the first worksheet while updating formula references, and saves the file as DeletedBlankRows.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data with some blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Data1");
            // Row 3 will be blank
            cells["A4"].PutValue("Data2");
            // Row 5 will be blank
            cells["A6"].PutValue("Data3");

            // Set up DeleteOptions with UpdateReference = true
            DeleteOptions options = new DeleteOptions
            {
                UpdateReference = true
            };

            // Delete all blank rows on the first worksheet using the options
            cells.DeleteBlankRows(options);

            // Save the workbook
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
