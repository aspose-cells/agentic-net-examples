// Title: Delete blank rows on the first worksheet with DeleteOptions.UpdateReference = true using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete blank rows on first sheet and keep formulas | DeleteBlankRows method with UpdateReference true example in .NET | How to remove empty rows from a worksheet while updating cell references using Aspose.Cells | C# code to delete completely blank rows in the first worksheet with DeleteOptions | Preserve formula references after deleting rows in Aspose.Cells workbook
// Tags: DeleteBlankRows DeleteOptions UpdateReference | first worksheet blank row removal Aspose.Cells | preserve formulas after row deletion .NET | Aspose.Cells remove empty rows C# | cell reference update on row deletion

using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, configures DeleteOptions to update cell references, deletes all fully blank rows on that sheet, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Configure delete options to update cell references after deletion
        DeleteOptions deleteOptions = new DeleteOptions();
        deleteOptions.UpdateReference = true;

        // Delete all completely blank rows in the first worksheet
        sheet.Cells.DeleteBlankRows(deleteOptions);

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
