// Title: Copy an existing worksheet by name to a new sheet in the same workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy worksheet by its name to a new sheet in the same workbook | How to duplicate an Excel worksheet while preserving formatting using Aspose.Cells .NET | Copy and rename a worksheet in Aspose.Cells without losing styles | Add a new worksheet and copy content from an existing sheet with Aspose.Cells C# | Preserve formulas and formatting when copying a worksheet using Aspose.Cells
// Tags: worksheet duplication within workbook Aspose.Cells | retain cell styles when copying Excel sheet C# | programmatic worksheet rename after copy Aspose.Cells | create new sheet and clone source worksheet Aspose.Cells | copy sheet content and settings Aspose.Cells

using Aspose.Cells;

// The example loads 'input.xlsx', retrieves the worksheet named 'Sheet1', adds a new worksheet, copies all content, formatting, formulas, and settings from the source to the new sheet, renames it to 'Sheet1_Copy', and saves the workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Specify the name of the worksheet to copy
        string sourceSheetName = "Sheet1";

        // Retrieve the source worksheet by name
        Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];

        // Add a new worksheet at the end of the workbook (this will be the target)
        int targetIndex = workbook.Worksheets.Add();

        // Get a reference to the newly added worksheet
        Worksheet targetSheet = workbook.Worksheets[targetIndex];

        // Copy the entire content, formatting, and settings from the source worksheet
        targetSheet.Copy(sourceSheet);

        // Optionally rename the copied worksheet to indicate it is a copy
        targetSheet.Name = sourceSheetName + "_Copy";

        // Save the workbook, preserving all copied content
        workbook.Save("output.xlsx");
    }
}
