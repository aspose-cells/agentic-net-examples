// Title: Export an Excel workbook to HTML while suppressing workbook properties with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSX file and saves it as HTML using Aspose.Cells, ensuring that workbook metadata is not included in the output. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to turn off the export of workbook properties during HTML conversion.
// Common Searches: how to save Excel as HTML without workbook metadata using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions disable workbook properties example | C# convert .xlsx to .html while omitting workbook properties | remove workbook properties from HTML output Aspose.Cells | Aspose.Cells HTML export settings to exclude metadata
// Tags: Aspose.Cells HtmlSaveOptions ExportWorkbookProperties false | HTML export without workbook metadata Aspose.Cells | C# convert XLSX to HTML Aspose.Cells | disable workbook properties in HTML output | Aspose.Cells HTML conversion options

// Create or load the workbook
// (Assuming the workbook is loaded from an existing file)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Configure HTML save options to omit workbook properties
Aspose.Cells.HtmlSaveOptions htmlOptions = new Aspose.Cells.HtmlSaveOptions(Aspose.Cells.SaveFormat.Html);
htmlOptions.ExportWorkbookProperties = false;   // Disable exporting of workbook properties

// Export the workbook to HTML
workbook.Save("output.html", htmlOptions);
