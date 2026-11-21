// Title: Load an Excel workbook (XLS or XLSX) into an Aspose.Cells Workbook using C# constructor or Load method
// Common Searches: C# code to open an existing XLSX file with Aspose.Cells Workbook constructor | How to load an Excel workbook from a file path using Aspose.Cells Load method in .NET | Aspose.Cells read Excel file from disk in C# without saving first | Example of initializing Aspose.Cells Workbook with an XLS file path in C# | Load Excel workbook into Aspose.Cells for further processing in .NET Core
// Tags: Aspose.Cells workbook constructor loading Excel file | Aspose.Cells Load method file path | C# load XLSX with Aspose.Cells | initialize Aspose.Cells workbook from local file | read Excel workbook using Aspose.Cells API

using Aspose.Cells;

// Demonstrates loading an Excel workbook (XLS or XLSX) in C# by passing the file path to the Aspose.Cells Workbook constructor, with an alternative example showing how to use the Workbook.Load method for the same operation.
class Program
{
    static void Main()
    {
        // Path to the Excel file (XLS or XLSX)
        string filePath = "input.xlsx";

        // Load the workbook using the constructor (preferred for direct loading)
        Workbook workbook = new Workbook(filePath);

        // Alternatively, you could use the Load method:
        // Workbook workbook = new Workbook();
        // workbook.Load(filePath);

        // At this point the workbook is loaded and ready for further processing.
    }
}
