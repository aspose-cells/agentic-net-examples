// Title: Load an Excel workbook and obtain a reference to cell B2 on Sheet1 using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file with Aspose.Cells, selects the worksheet named "Sheet1", and retrieves the Cell object for address B2. | Show how to access a worksheet by its name and get a specific cell reference in Aspose.Cells, then read the cell's value. | Demonstrate retrieving a cell by address after opening a workbook with Aspose.Cells, including handling of workbook and worksheet objects.
// Common Searches: aspnet aspose.cells get cell B2 from specific worksheet | c# read value of B2 in Excel using Aspose.Cells library | how to reference a cell by address after opening workbook with Aspose.Cells | select worksheet by name and access cell in Aspose.Cells C# example | retrieve cell value from Sheet1 using Aspose.Cells .NET API
// Tags: open workbook Aspose.Cells C# | worksheet name selection Aspose.Cells | cell address access Aspose.Cells | read Excel cell value .NET | Aspose.Cells reference specific cell

using Aspose.Cells;

// The example opens an existing .xlsx file with Aspose.Cells, selects the worksheet named "Sheet1", obtains a reference to cell B2, and reads its value.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        // Access the desired worksheet by name (or use index: workbook.Worksheets[0])
        Worksheet worksheet = workbook.Worksheets["Sheet1"]; // worksheet access

        // Obtain a reference to the target cell (e.g., cell B2)
        Cell targetCell = worksheet.Cells["B2"]; // cell reference

        // Example usage: read the cell's value
        var cellValue = targetCell.Value;
    }
}
