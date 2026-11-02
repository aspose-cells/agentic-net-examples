// Title: Add a cell containing a formula to the Formula Watch Window with Worksheet.CellWatches.Add in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add cell A1 to Formula Watch Window after setting formula | How to use Worksheet.CellWatches.Add to monitor a formula in a workbook | Aspose.Cells .NET example for adding a formula watch for a specific cell | Programmatically add a cell to the Formula Watch list in Aspose.Cells | Watch a SUM formula in Excel using Aspose.Cells CellWatches API | Add formula watch for range B1:B10 using Aspose.Cells C#
// Tags: Worksheet.CellWatches.Add API usage | add formula watch cell Aspose.Cells | monitor Excel formula with Aspose.Cells C# | Formula Watch Window example .NET | Aspose.Cells add cell watch to workbook | C# set cell formula and watch list

using Aspose.Cells;
using System;

// The sample creates a new workbook, assigns a SUM formula to cell A1 on the first worksheet, adds A1 to the worksheet's Formula Watch Window using Worksheet.CellWatches.Add, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set a formula in cell A1
        Cell cell = sheet.Cells["A1"];
        cell.Formula = "=SUM(B1:B10)";

        // Add cell A1 to the Formula Watch Window
        sheet.CellWatches.Add("A1");

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
