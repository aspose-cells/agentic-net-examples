// Title: How to load an Excel workbook from a file and get the first worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load workbook from local file path and read first sheet | C# code to open input.xlsx with Aspose.Cells and access worksheet index 0 | Example of retrieving the first worksheet after loading an Excel file using Aspose.Cells .NET | Load Excel file and get first worksheet programmatically with Aspose.Cells library | Aspose.Cells workbook initialization from file and first worksheet extraction in C#
// Tags: load workbook from file Aspose.Cells C# | access first worksheet by index Aspose.Cells | initialize Workbook with file path .NET | retrieve worksheet from Worksheets collection Aspose.Cells | open Excel file programmatically Aspose.Cells

using Aspose.Cells;

// Path to the Excel file
string filePath = "input.xlsx";

// Load the workbook from the specified file
Workbook workbook = new Workbook(filePath);

// Access the first worksheet (index 0)
Worksheet firstWorksheet = workbook.Worksheets[0];

// You can now work with 'firstWorksheet' as needed.
