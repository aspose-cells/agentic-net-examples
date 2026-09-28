// Title: Open an XLSB binary workbook from a file path using Aspose.Cells for .NET
// AI Prompts: Use the Aspose.Cells Workbook constructor to load an .xlsb file from a given path and retrieve its first worksheet in C#. | Create a C# program that opens a binary Excel workbook (.xlsb) by passing the file location to Aspose.Cells without specifying the format. | Demonstrate automatic format detection when initializing a Workbook object with a path to an .xlsb file in .NET.
// Common Searches: asp.net how to load .xlsb workbook with Aspose.Cells without setting format | c# Aspose.Cells open binary Excel file from file system | example code for automatic file format detection in Aspose.Cells Workbook constructor | read first worksheet of an .xlsb file using Aspose.Cells in C# | load Excel binary workbook path Aspose.Cells .NET tutorial
// Tags: load xlsb workbook Aspose.Cells | automatic format detection Aspose.Cells | C# open binary Excel file | Workbook constructor file path | retrieve first worksheet C# | Aspose.Cells .xlsb file handling

using Aspose.Cells;

// The example shows how to instantiate a Workbook object by providing the file path of an XLSB binary workbook; Aspose.Cells automatically detects the format, and the code then accesses the first worksheet for further processing.
class Program
{
    static void Main()
    {
        // Path to the XLSB binary workbook
        string filePath = @"C:\Path\To\YourWorkbook.xlsb";

        // The Workbook constructor automatically detects the XLSB format and loads the file
        Workbook workbook = new Workbook(filePath);

        // Example: access the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];

        // Additional processing can be performed here
    }
}
