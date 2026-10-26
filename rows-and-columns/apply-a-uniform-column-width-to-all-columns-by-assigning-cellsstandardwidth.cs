// Title: Set a uniform column width for every column in an Excel worksheet using Cells.StandardWidth with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to apply the same column width to all columns in a worksheet | using Cells.StandardWidth to set default column width for an entire sheet in Aspose.Cells | Aspose.Cells set uniform column width across all columns programmatically in .NET
// Tags: Cells.StandardWidth property usage | uniform column width Aspose.Cells C# | apply column width to entire worksheet Aspose.Cells | Excel column width setting Aspose.Cells .NET | standard column width property Aspose.Cells

using Aspose.Cells;
using System;

// The example creates or loads a workbook, accesses the first worksheet, assigns Cells.StandardWidth = 12 to give every column the same width measured in characters, and saves the file as UniformColumnWidth.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // default workbook with one worksheet

        // Apply a uniform column width to all columns in the first worksheet
        // StandardWidth is measured in characters of the default font
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.StandardWidth = 12; // set width to 12 characters

        // Save the workbook
        workbook.Save("UniformColumnWidth.xlsx", SaveFormat.Xlsx);
    }
}
