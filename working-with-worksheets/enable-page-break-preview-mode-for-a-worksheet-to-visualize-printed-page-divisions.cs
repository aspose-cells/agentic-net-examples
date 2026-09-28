// Title: How to enable page break preview for a worksheet using Aspose.Cells in C#
// AI Prompts: Assign true to worksheet.IsPageBreakPreview and save the workbook as an XLSX file with Aspose.Cells. | Programmatically toggle page break preview on a selected worksheet in a .NET project and export the result. | Create a new workbook, turn on page break preview for the first sheet, and write the file to disk using Aspose.Cells.
// Common Searches: Aspose.Cells C# enable page break preview for a sheet | How to view printed page boundaries in Excel with Aspose.Cells .NET | Set IsPageBreakPreview property programmatically using Aspose.Cells | C# example to turn on page break preview mode in an Excel workbook
// Tags: Aspose.Cells page break preview activation | C# worksheet page break preview setting | Excel printed page boundaries visualization | Aspose.Cells workbook save with preview enabled | Aspose.Cells page break view activation

using Aspose.Cells;

// Create a new workbook (or load an existing one)
// Replace this with the provided create/load rule if available
Workbook workbook = new Workbook();

// Access the first worksheet (or any specific worksheet)
Worksheet worksheet = workbook.Worksheets[0];

// Enable Page Break Preview mode to visualize printed page divisions
worksheet.IsPageBreakPreview = true;

// Save the workbook (replace with the provided save rule if available)
workbook.Save("PageBreakPreview.xlsx");
