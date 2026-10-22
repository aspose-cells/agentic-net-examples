// Title: Load an XLSX workbook from a local file using Aspose.Cells Workbook class in C#
// AI Prompts: Instantiate a Workbook object by supplying the file path of an existing .xlsx document to the Aspose.Cells constructor. | Open an Excel workbook stored on disk and obtain a ready‑to‑use Workbook instance for further processing with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# how to open an existing .xlsx file from a file path | Load Excel workbook from disk using Aspose.Cells Workbook constructor example | Read .xlsx file into Aspose.Cells Workbook object in a .NET console application
// Tags: load xlsx workbook from local path Aspose.Cells | initialize Workbook from file path C# | open Excel workbook via Aspose.Cells constructor | read spreadsheet file using Aspose.Cells C#

using Aspose.Cells;

// Demonstrates loading an XLSX file (e.g., "input.xlsx") from the file system into an Aspose.Cells Workbook object for further manipulation in a C# application.
class Program
{
    static void Main()
    {
        // Path to the XLSX file on disk
        string filePath = "input.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(filePath);

        // The workbook is now loaded and can be manipulated as needed
    }
}
