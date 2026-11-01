// Title: Load an existing .xlsx file into an Aspose.Cells Workbook object in C# for further manipulation
// AI Prompts: Generate C# code that uses Aspose.Cells to open a local .xlsx file and create a Workbook instance. | Show how to load an Excel workbook into memory with Aspose.Cells so I can access its worksheets programmatically. | Provide a minimal example of initializing a Workbook from a file path and retrieving the first worksheet using Aspose.Cells in .NET.
// Common Searches: how to read an .xlsx file into a Workbook object using Aspose.Cells C# | Aspose.Cells example for loading Excel workbook from disk in .NET | C# code to open Excel file with Aspose.Cells and access worksheets | initialize Aspose.Cells Workbook from file path for manipulation
// Tags: Aspose.Cells load workbook from .xlsx | C# instantiate Workbook from file path | Aspose.Cells read Excel file into memory | access first worksheet Aspose.Cells C# | load Excel workbook for manipulation Aspose.Cells

using Aspose.Cells;

// Loads the "input.xlsx" file into an Aspose.Cells Workbook object in memory, enabling further operations such as accessing worksheets.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // The workbook is now in memory and can be manipulated further
        // Example: Access the first worksheet
        // Worksheet sheet = workbook.Worksheets[0];
    }
}
