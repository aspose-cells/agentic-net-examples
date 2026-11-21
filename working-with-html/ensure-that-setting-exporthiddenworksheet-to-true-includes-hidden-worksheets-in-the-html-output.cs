// Title: Export hidden Excel worksheets to HTML using Aspose.Cells HtmlSaveOptions in C#
// Common Searches: Aspose.Cells C# export hidden Excel sheets to HTML | HtmlSaveOptions ExportHiddenWorksheet property example | Save workbook as HTML including hidden worksheets Aspose.Cells | C# generate HTML from Excel with hidden sheets using Aspose.Cells | How to include hidden worksheets when converting Excel to HTML in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet | C# HTML conversion of hidden Excel sheets | Aspose.Cells save workbook as HTML with hidden sheets | include hidden worksheets in HTML output Aspose.Cells | Aspose.Cells hidden worksheet export setting

// Load an existing workbook (replace with your actual file path)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Configure HTML save options to include hidden worksheets
Aspose.Cells.HtmlSaveOptions htmlOptions = new Aspose.Cells.HtmlSaveOptions();
htmlOptions.ExportHiddenWorksheet = true;   // Ensure hidden sheets are exported

// Save the workbook as HTML (replace with your desired output path)
workbook.Save("output.html", htmlOptions);
