// Title: Save a modified Excel workbook as XLSB while preserving external data connections using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# how to save an Excel file as XLSB without breaking external connections | convert XLSX to binary XLSB format while keeping data source links in .NET | C# example for retaining external data connections when exporting to XLSB with Aspose.Cells | saving a workbook as XLSB and preserving connection strings using Aspose.Cells API | modify a cell and export to XLSB while maintaining Excel data connections in C#
// Tags: XLSB export with external connections Aspose.Cells | Aspose.Cells SaveFormat.Xlsb usage | C# preserve data connections during Excel binary save | convert Excel to binary workbook retaining connections | modify workbook cell before XLSB conversion Aspose.Cells

using Aspose.Cells;
using System;

// Loads an existing workbook, updates cell A1, and saves it as an XLSB file using Aspose.Cells. The SaveFormat.Xlsb operation automatically retains any external connection settings, ensuring data links remain intact after conversion.
class Program
{
    static void Main()
    {
        // Load the existing workbook (any supported format)
        Workbook workbook = new Workbook("input.xlsx");

        // Perform modifications as needed (example: change a cell value)
        workbook.Worksheets[0].Cells["A1"].PutValue("Modified");

        // External connections are retained automatically when saving.
        // Save the workbook as an XLSB file.
        workbook.Save("output.xlsb", SaveFormat.Xlsb);
    }
}
