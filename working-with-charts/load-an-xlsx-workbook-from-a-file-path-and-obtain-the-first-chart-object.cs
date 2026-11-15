// Title: Load an XLSX workbook from a file path and obtain the first chart object in C# with Aspose.Cells
// Common Searches: C# Aspose.Cells load workbook from file path and get first chart | How to access the first chart in an Excel worksheet using Aspose.Cells .NET | Retrieve chart collection from XLSX file with Aspose.Cells in C# | Aspose.Cells example for reading charts from a workbook | Get chart object from worksheet using Aspose.Cells API
// Tags: load xlsx workbook Aspose.Cells C# | first chart extraction Aspose.Cells | worksheet chart collection access C# | Aspose.Cells read chart objects from Excel | retrieve chart from XLSX using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;

// Load the workbook from the specified file path
string filePath = @"C:\Path\To\YourWorkbook.xlsx";
Workbook workbook = new Workbook(filePath);

// Access the first worksheet (index 0)
Worksheet worksheet = workbook.Worksheets[0];

// Retrieve the first chart object from the worksheet
Chart firstChart = null;
if (worksheet.Charts.Count > 0)
{
    firstChart = worksheet.Charts[0];
}

// At this point 'firstChart' holds the first chart in the worksheet (or null if none)
