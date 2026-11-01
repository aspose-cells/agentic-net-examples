// Title: How to assign an absolute address to a named range in Aspose.Cells for .NET to keep it from shifting
// Common Searches: Aspose.Cells C# set named range RefersTo to absolute address | prevent named range from moving after inserting rows or columns in Aspose.Cells | create fixed reference named range using worksheet name in Aspose.Cells | absolute address syntax for named ranges in Aspose.Cells .NET | Aspose.Cells example adding named range with $A$1:$B$10 reference
// Tags: Aspose.Cells named range absolute RefersTo | fixed named range address Aspose.Cells | C# Aspose.Cells set RefersTo property | absolute cell reference in named range Aspose.Cells | prevent named range shift Aspose.Cells

using Aspose.Cells;

// Shows how to build a fully qualified absolute address, add a named range to the workbook's Names collection, assign the RefersTo property to that address, and save the file so the range remains unchanged when rows or columns are inserted or deleted.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet (default sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Define an absolute address for the named range (e.g., $A$1:$B$10)
        // Using the worksheet name ensures the reference is fully qualified.
        string absoluteAddress = $"{sheet.Name}!$A$1:$B$10";

        // Add a new named range to the workbook's Names collection
        int nameIndex = workbook.Worksheets.Names.Add("MyAbsoluteRange");

        // Set the RefersTo property using the absolute address.
        // This prevents the range from shifting when rows/columns are inserted or deleted.
        workbook.Worksheets.Names[nameIndex].RefersTo = absoluteAddress;

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
