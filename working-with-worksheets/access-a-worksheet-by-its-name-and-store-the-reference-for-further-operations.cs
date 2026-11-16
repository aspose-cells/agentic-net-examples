// Title: Access a worksheet by its name and obtain a Worksheet object using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# get worksheet object by sheet name from existing workbook | How to reference a specific worksheet by name in Aspose.Cells for .NET | C# code to load workbook and select sheet named MySheet with Aspose.Cells | Retrieve worksheet reference using Aspose.Cells Worksheets indexer in C# | Select worksheet by name in Aspose.Cells without using index number
// Tags: worksheet name lookup Aspose.Cells | Aspose.Cells Worksheets indexer usage | load workbook and select sheet by name C# | C# get specific worksheet object | Aspose.Cells reference specific sheet

using Aspose.Cells;

// Loads an existing workbook (e.g., input.xlsx), accesses the worksheet named "MySheet" via the Worksheets collection indexer, and stores the result in a Worksheet variable for further manipulation.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the file path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet by its name and store the reference
        Worksheet worksheet = workbook.Worksheets["MySheet"]; // replace "MySheet" with the actual sheet name

        // The 'worksheet' variable can now be used for further operations
    }
}
