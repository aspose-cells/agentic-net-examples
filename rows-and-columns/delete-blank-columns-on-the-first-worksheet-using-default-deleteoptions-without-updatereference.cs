// Title: Delete blank columns from the first worksheet in an Excel file using Aspose.Cells for .NET with DeleteOptions (UpdateReference = false)
// Common Searches: Aspose.Cells C# delete empty columns without updating formulas | How to remove blank columns from the first sheet using DeleteOptions in Aspose.Cells | C# code to delete blank columns in an .xlsx workbook with UpdateReference set to false | Aspose.Cells DeleteBlankColumns method example for first worksheet | Remove empty columns from Excel using Aspose.Cells default DeleteOptions
// Tags: Aspose.Cells DeleteBlankColumns API | DeleteOptions UpdateReference false | remove blank columns first worksheet | C# Excel column cleanup Aspose.Cells | default DeleteOptions usage Aspose.Cells

using Aspose.Cells;

// The example loads an .xlsx workbook, accesses the first worksheet, creates a DeleteOptions object with UpdateReference disabled, calls Cells.DeleteBlankColumns to remove all empty columns, and saves the modified workbook.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Configure DeleteOptions without updating references
        DeleteOptions deleteOptions = new DeleteOptions();
        deleteOptions.UpdateReference = false;

        // Delete all blank columns on the first worksheet
        sheet.Cells.DeleteBlankColumns(deleteOptions);

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
