// Title: Set a default column width of 15 characters for the entire worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Assign 15 to worksheet.Cells.StandardWidth in a C# Aspose.Cells workbook and then save the file. | Create a new workbook, set a uniform column width of 15 characters for all columns, and export to Result.xlsx using Aspose.Cells. | Configure the worksheet's default column width to 15 characters before writing any data with Aspose.Cells in .NET.
// Common Searches: Aspose.Cells C# how to set default column width for every column in a worksheet | Set worksheet StandardWidth to 15 characters using Aspose.Cells .NET | C# example for applying a uniform column width across an entire Excel sheet with Aspose.Cells | Change default column width for a new workbook before saving with Aspose.Cells | Aspose.Cells StandardWidth property tutorial in C#
// Tags: Aspose.Cells worksheet default column width | C# StandardWidth property usage | set column width for entire sheet Aspose.Cells | Excel workbook column width setting Aspose.Cells | Aspose.Cells uniform column width C#

using Aspose.Cells;
using System;

// The code creates a new workbook, accesses the first worksheet, sets its default column width to 15 characters via the StandardWidth property, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the default column width to 15 characters for the entire sheet
        worksheet.Cells.StandardWidth = 15;

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
