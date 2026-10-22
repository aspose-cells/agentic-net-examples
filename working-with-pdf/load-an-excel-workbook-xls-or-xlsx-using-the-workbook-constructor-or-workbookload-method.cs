// Title: How to load an XLS or XLSX workbook from a file with Aspose.Cells in C# using the Workbook constructor or Load method
// AI Prompts: Write C# code that creates a Workbook instance by passing the path of an .xlsx file to the Aspose.Cells Workbook constructor. | Show a C# example that initializes an empty Workbook object and then calls the Load method to open an Excel file from disk. | Provide a single Aspose.Cells snippet that determines the file extension and loads either a .xls or .xlsx file into a Workbook.
// Common Searches: asp.net core load excel file using aspose.cells workbook constructor | c# load .xls file with aspose.cells without saving first | difference between Workbook constructor and Load method in Aspose.Cells | how to open an existing workbook from a path in Aspose.Cells C# | load excel workbook from file path using Aspose.Cells in .NET 6
// Tags: Aspose.Cells load workbook from file path | Workbook constructor for .xlsx in C# | Workbook.Load method for .xls files | .NET read Excel file using Aspose.Cells | initialize Aspose.Cells workbook without saving

using Aspose.Cells;

// // Loads an Excel workbook from 'input.xlsx' using the Aspose.Cells Workbook constructor; alternative Load method shown.
class Program
{
    static void Main()
    {
        // Path to the Excel file (XLS or XLSX)
        string filePath = "input.xlsx";

        // Load the workbook using the constructor
        Workbook workbook = new Workbook(filePath);

        // Alternatively, you can load using the Load method:
        // Workbook workbook = new Workbook();
        // workbook.Load(filePath);
    }
}
