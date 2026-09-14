// Title: How to set worksheet page zoom to 80% with Aspose.Cells for .NET and export to PDF
// AI Prompts: Apply an 80% zoom to the active sheet's page layout and generate a PDF file using Aspose.Cells in C#. | Adjust the zoom level of several worksheets before creating a combined PDF document in a .NET application. | Change the page scaling factor for PDF export while preserving existing cell data with Aspose.Cells.
// Common Searches: aspnet set page zoom 80% before exporting Excel to PDF with Aspose.Cells | c# Aspose.Cells change worksheet PageSetup.Zoom for PDF output | how to adjust PDF scaling of Excel sheet using Aspose.Cells .NET | export workbook to PDF with custom zoom level using Aspose.Cells C# example
// Tags: worksheet zoom setting Aspose.Cells | custom PDF scaling using Aspose.Cells | PageSetup.Zoom API C# | export Excel to PDF with adjusted layout | modify page setup prior to PDF generation

using Aspose.Cells;
using System;

// Creates a new workbook, sets the first worksheet's PageSetup.Zoom to 80%, adds optional sample data, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the page zoom to 80%
        sheet.PageSetup.Zoom = 80;

        // (Optional) Add some sample data to the worksheet
        sheet.Cells["A1"].PutValue("Sample data for PDF export");

        // Export the workbook to PDF
        workbook.Save("Result.pdf", SaveFormat.Pdf);
    }
}
