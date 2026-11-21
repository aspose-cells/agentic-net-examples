// Title: Load a local XLSX workbook with Aspose.Cells Workbook constructor in C#
// Common Searches: C# Aspose.Cells load Excel workbook from absolute file path | How to instantiate Workbook with a local .xlsx file using Aspose.Cells | Reading first worksheet after opening XLSX via Aspose.Cells in .NET | Aspose.Cells default constructor example for opening Excel files in C# | Open Excel file stored on disk with Aspose.Cells C# code
// Tags: Aspose.Cells Workbook constructor file path | C# open local XLSX with Aspose.Cells | Aspose.Cells retrieve first worksheet after load | Aspose.Cells read Excel file from disk | initialize Aspose.Cells Workbook from local file

using Aspose.Cells;

// Demonstrates how to specify a local XLSX file path, create a Workbook instance using Aspose.Cells' default constructor, and access the first worksheet for further processing in C#.
class Program
{
    static void Main()
    {
        // Specify the full path to the XLSX file
        string filePath = @"C:\Path\To\Your\File.xlsx";

        // Open the workbook using the default Workbook constructor
        Workbook workbook = new Workbook(filePath);

        // Example: access the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];

        // Additional processing can be done here
    }
}
