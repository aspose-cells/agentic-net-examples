// Title: Open a spreadsheet file of any supported format by providing only the file path to the Aspose.Cells Workbook constructor in C#
// Common Searches: C# Aspose.Cells load workbook from file path without specifying format | How to let Aspose.Cells automatically detect xls, xlsx, csv when opening | Aspose.Cells Workbook constructor example using only the file name | Open Excel or CSV file with Aspose.Cells by providing just the path | Automatic spreadsheet format detection in .NET using Aspose.Cells
// Tags: automatic workbook format detection Aspose.Cells | Workbook constructor file path only | load spreadsheet without format argument C# | open multiple spreadsheet formats Aspose.Cells | Aspose.Cells load workbook by path

using Aspose.Cells;

// The code creates a Workbook by passing only the file path, letting Aspose.Cells automatically determine the workbook format (e.g., .xls, .xlsx, .csv) and then accesses the first worksheet.
class Program
{
    static void Main()
    {
        // Provide only the file path to the Workbook constructor.
        // Aspose.Cells automatically detects the workbook format (xls, xlsx, csv, etc.).
        string filePath = @"C:\Data\SampleWorkbook.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Example: Access the first worksheet to verify loading succeeded.
        Worksheet sheet = workbook.Worksheets[0];
        // Perform any additional operations here...

        // Optionally, save the workbook in a different format.
        // workbook.Save(@"C:\Data\ConvertedWorkbook.pdf", SaveFormat.Pdf);
    }
}
