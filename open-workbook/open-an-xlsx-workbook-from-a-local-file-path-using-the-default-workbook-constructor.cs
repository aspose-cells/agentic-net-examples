// Title: Load an XLSX workbook from a local file path using Aspose.Cells Workbook constructor in C#
// Common Searches: C# Aspose.Cells load workbook from absolute file path | How to instantiate Workbook with a local XLSX file in .NET | Aspose.Cells open Excel file from disk using default constructor | Reading an existing XLSX file with Aspose.Cells in a C# console app | Load Excel workbook from C:\Data\SampleWorkbook.xlsx using Aspose.Cells
// Tags: Aspose.Cells workbook constructor with file name | load xlsx workbook using default constructor | initialize Aspose.Cells workbook from local file | open existing Excel workbook in C# Aspose.Cells | read Excel file from disk with Aspose.Cells

using Aspose.Cells;

// The example shows how to specify a full file path to an XLSX workbook and create a new Aspose.Cells Workbook instance by passing that path to the default constructor, thereby loading the workbook for further processing in C#.
class Program
{
    static void Main()
    {
        // Specify the full path to the XLSX file
        string filePath = @"C:\Data\SampleWorkbook.xlsx";

        // Open the workbook using the default Workbook constructor
        Workbook workbook = new Workbook(filePath);

        // The workbook is now loaded and can be used for further processing
    }
}
