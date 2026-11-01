// Title: How to load an .xlsx workbook from disk and retrieve the first worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a local .xlsx file with Aspose.Cells and returns the Worksheet object at index 0. | Create a snippet that reads the string value of cell A1 from the first worksheet after loading a workbook with Aspose.Cells. | Demonstrate accessing a worksheet by its zero‑based index in Aspose.Cells and extracting a cell value.
// Common Searches: Aspose.Cells C# load workbook from file and get first sheet | Read cell A1 from the first worksheet using Aspose.Cells .NET | How to access worksheet by index in Aspose.Cells C# example
// Tags: load workbook from .xlsx Aspose.Cells C# | access first worksheet by index Aspose.Cells | read cell value A1 Aspose.Cells | Aspose.Cells workbook initialization C# | worksheet collection indexing Aspose.Cells

using Aspose.Cells;

// // Loads 'input.xlsx' into a Workbook, accesses the first worksheet (index 0), and reads the string value of cell A1.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example: read the value of cell A1
        string cellValue = firstWorksheet.Cells["A1"].StringValue;
    }
}
