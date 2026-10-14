// Title: How to retrieve a worksheet by its name and keep a reference using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an Excel file with Aspose.Cells and obtains a Worksheet object by specifying the sheet name. | Show how to assign the Worksheet returned from workbook.Worksheets["SheetName"] to a variable for later cell operations.
// Common Searches: Aspose.Cells C# get worksheet object by sheet name from existing workbook | C# code to access a specific sheet in an Excel file using Aspose.Cells Worksheets collection | How to reference a worksheet for further manipulation after loading a workbook with Aspose.Cells | Retrieve worksheet named "Sheet1" in .NET using Aspose.Cells API
// Tags: Aspose.Cells retrieve worksheet by name | C# store Worksheet reference | Aspose.Cells load workbook from xlsx | worksheet object for further manipulation Aspose.Cells | access specific sheet Aspose.Cells .NET

using Aspose.Cells;

// The example loads an existing workbook (input.xlsx) with Aspose.Cells, accesses the worksheet named "Sheet1" via the Worksheets collection, and stores the Worksheet object in a variable for subsequent operations such as reading or writing cell values.
class WorksheetAccessExample
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet by its name and store the reference
        // Ensure the sheet name exists in the workbook
        Worksheet sheet = workbook.Worksheets["Sheet1"];

        // The 'sheet' variable can now be used for further operations
        // Example: read a cell value
        // var value = sheet.Cells["A1"].StringValue;
    }
}
