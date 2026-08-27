// Title: Delete blank rows from the first worksheet of a C# Aspose.Cells workbook using the default DeleteBlankRows method (no UpdateReference)
// AI Prompts: Use Aspose.Cells in C# to remove all empty rows from the first sheet without adjusting cell references. | Show how to call Cells.DeleteBlankRows with default options to clean up a workbook created in code. | Generate C# code that deletes blank rows on the first worksheet and saves the result as an .xlsx file.
// Common Searches: Aspose.Cells C# delete blank rows first worksheet default options | How to remove empty rows from an Excel sheet using Aspose.Cells without updating formulas | DeleteBlankRows method example without UpdateReference in .NET | C# code to clean up blank rows in a newly created workbook using Aspose.Cells | Aspose.Cells remove blank rows from worksheet and keep cell references unchanged
// Tags: Aspose.Cells DeleteBlankRows default | C# remove empty rows Excel worksheet | DeleteBlankRows without UpdateReference | blank row cleanup Aspose.Cells | first worksheet row deletion C#

using System;
using Aspose.Cells;

namespace DeleteBlankRowsExample
{
    // The example creates a new workbook, adds sample data with intentional blank rows, calls Cells.DeleteBlankRows() on the first worksheet using the default DeleteOptions (which do not update references), and saves the cleaned workbook as DeletedBlankRows.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with intentional blank rows
            cells["A1"].PutValue("Header");
            cells["A2"].PutValue("Row1");
            // Row 3 will be blank
            cells["A4"].PutValue("Row2");
            // Row 5 will be blank
            cells["A6"].PutValue("Row3");

            // Delete all blank rows using the default method (no UpdateReference)
            cells.DeleteBlankRows();

            // Save the result
            workbook.Save("DeletedBlankRows.xlsx", SaveFormat.Xlsx);
        }
    }
}
