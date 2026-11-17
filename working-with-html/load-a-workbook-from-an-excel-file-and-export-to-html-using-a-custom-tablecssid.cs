// Title: Export Excel to HTML with a custom TableCssId using Aspose.Cells for .NET (C#)
// Common Searches: how to assign a custom TableCssId when converting Excel to HTML with Aspose.Cells in C# | Aspose.Cells C# export workbook to HTML with specific table id attribute | set TableCssId property in HtmlSaveOptions for Excel to HTML conversion | C# code example for saving Excel as HTML with custom table CSS identifier using Aspose.Cells | generate HTML from Excel with a custom table id using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | Excel to HTML conversion C# Aspose | custom table identifier Aspose.Cells | HTML save options for Excel workbook .NET | TableCssId configuration Aspose.Cells

using System;
using Aspose.Cells;

// The sample loads an Excel workbook (input.xlsx), sets HtmlSaveOptions.TableCssId to "myCustomTableId", and saves the workbook as an HTML file (output.html) using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options with a custom CSS ID for the generated table
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.TableCssId = "myCustomTableId";

        // Export the workbook to an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
