// Title: Load an Excel workbook (XLS or XLSX) from a file path using Aspose.Cells Workbook constructor or Workbook.Load in C#
// AI Prompts: Write C# code that creates a Workbook instance from a .xlsx file path using the Aspose.Cells constructor. | Demonstrate how to call Workbook.Load to open an existing .xls file with Aspose.Cells, including basic error handling. | Show a C# snippet that decides at runtime whether to use the constructor or Load method for opening Excel files with Aspose.Cells.
// Common Searches: aspocells c# open existing xls file from disk | difference between Workbook constructor and Load method in Aspose.Cells | how to read an Excel workbook using Aspose.Cells without saving changes | c# load excel workbook from path using Aspose.Cells library | example of loading both .xls and .xlsx with Aspose.Cells in .NET
// Tags: initialize workbook using file location Aspose.Cells C# | invoke Workbook.Load for Excel import Aspose.Cells | open XLSX workbook programmatically Aspose.Cells | read Excel file from local storage Aspose.Cells | error handling for unsupported Excel format Aspose.Cells | select loading strategy for Excel import Aspose.Cells

using Aspose.Cells;

// Demonstrates loading an Excel workbook (XLS or XLSX) from a specified file path using Aspose.Cells' Workbook constructor, with an alternative approach shown using the Workbook.Load method.
class Program
{
    static void Main()
    {
        // Path to the Excel file (XLS or XLSX)
        string filePath = "input.xlsx";

        // Load the workbook using the constructor (preferred for simplicity)
        Workbook workbook = new Workbook(filePath);

        // Alternative way: load using the Load method
        // Workbook workbook = new Workbook();
        // workbook.Load(filePath);
    }
}
