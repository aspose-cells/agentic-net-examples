// Title: Export each worksheet to its own CSS file when saving an Excel workbook as HTML using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file and saves it as HTML, creating a separate CSS file for every worksheet with Aspose.Cells HtmlSaveOptions. | Show how to enable the ExportWorksheetCSSSeparately property in HtmlSaveOptions to produce per‑sheet stylesheet files in a .NET application. | Adapt an existing Aspose.Cells HTML export routine so that each worksheet generates its own CSS file instead of a single combined stylesheet.
// Common Searches: Aspose.Cells how to create separate CSS files for each worksheet when converting Excel to HTML in C# | C# HtmlSaveOptions ExportWorksheetCSSSeparately example code | Save Excel workbook as HTML with per‑sheet stylesheet using Aspose.Cells .NET | Generate distinct CSS per worksheet during Excel to HTML conversion with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportWorksheetCSSSeparately | C# Excel to HTML per worksheet CSS | separate stylesheet generation Aspose.Cells | HTML export distinct CSS per sheet | Aspose.Cells workbook save as HTML individual CSS files

// Load an existing workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Configure HTML save options to export a separate CSS file for each worksheet
Aspose.Cells.HtmlSaveOptions htmlOptions = new Aspose.Cells.HtmlSaveOptions();
htmlOptions.ExportWorksheetCSSSeparately = true;

// Save the workbook as HTML using the configured options
workbook.Save("output.html", htmlOptions);
