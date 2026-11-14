// Title: Load an Excel workbook from a file and retrieve the first worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load workbook from a file and get the first worksheet | How to read the name of the first worksheet in an Excel file using Aspose.Cells | Open input.xlsx with Aspose.Cells and access worksheet at index 0 in .NET | C# example for loading an Excel workbook and retrieving the first sheet via Aspose.Cells | Retrieve first worksheet object from a workbook using Aspose.Cells for .NET
// Tags: initialize workbook object from file Aspose.Cells | access worksheet by zero-based index Aspose.Cells | read worksheet name from Excel workbook .NET | open Excel file with Aspose.Cells API | worksheet collection indexing Aspose.Cells

using Aspose.Cells;

// Loads the Excel workbook located at "input.xlsx" with Aspose.Cells, obtains the first worksheet (index 0), and prints its name to the console.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet firstWorksheet = workbook.Worksheets[0];

        // Example usage: display the name of the first worksheet
        System.Console.WriteLine("First worksheet name: " + firstWorksheet.Name);
    }
}
