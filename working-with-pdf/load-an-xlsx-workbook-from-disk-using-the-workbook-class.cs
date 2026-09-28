// Title: Load an XLSX workbook from a file path with Aspose.Cells Workbook in C#
// AI Prompts: Write C# code that uses Aspose.Cells Workbook to open a .xlsx file from a given path and retrieve the first worksheet. | Show how to instantiate a Workbook object with a file name, then access a specific sheet using Aspose.Cells for .NET.
// Common Searches: aspnet load xlsx file using Aspose.Cells Workbook constructor | c# read Excel workbook from disk with Aspose.Cells and get first sheet | example of opening an .xlsx file with Aspose.Cells in a console application | how to initialize Aspose.Cells Workbook with a file path in C# | retrieve worksheet after loading workbook using Aspose.Cells .NET
// Tags: load workbook from file Aspose.Cells C# | read xlsx file Aspose.Cells .NET | access first worksheet Aspose.Cells | initialize Workbook with file path | open Excel workbook Aspose.Cells C#

using Aspose.Cells;

// The example shows how to create an Aspose.Cells Workbook object by providing the path to an .xlsx file, then optionally obtain the first worksheet for further processing.
class Program
{
    static void Main()
    {
        // Path to the XLSX file to load
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // Example: access the first worksheet (optional)
        Worksheet firstSheet = workbook.Worksheets[0];
    }
}
