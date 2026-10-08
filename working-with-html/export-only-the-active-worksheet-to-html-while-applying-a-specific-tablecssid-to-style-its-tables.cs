// Title: Export the active worksheet to HTML and assign a custom TableCssId with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, sets HtmlSaveOptions.ExportActiveWorksheetOnly to true, assigns a TableCssId value, and saves the result as an HTML file using Aspose.Cells. | Show how to apply a custom CSS ID to the tables produced by Aspose.Cells when exporting only the active worksheet to HTML.
// Common Searches: How to export only the current sheet to HTML with a specific table CSS ID using Aspose.Cells in C# | Aspose.Cells C# HtmlSaveOptions TableCssId example for active worksheet | Saving an Excel worksheet as HTML with a custom table identifier in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | Aspose.Cells TableCssId custom HTML tables | C# export active worksheet to HTML Aspose | custom CSS ID for Aspose.Cells HTML output | save Excel sheet as HTML with Aspose.Cells

// Load the workbook from a file (replace with your actual file path)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Configure HTML save options
var htmlOptions = new Aspose.Cells.HtmlSaveOptions(Aspose.Cells.SaveFormat.Html)
{
    // Export only the active worksheet
    ExportActiveWorksheetOnly = true,
    
    // Assign a CSS ID to the generated HTML tables for custom styling
    TableCssId = "MyCustomTable"
};

// Save the active worksheet as HTML using the configured options
workbook.Save("output.html", htmlOptions);
