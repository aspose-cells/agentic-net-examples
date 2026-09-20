// Title: Define a dynamic print area in an Aspose.Cells worksheet that grows with added rows and columns using C#
// AI Prompts: Generate C# code that creates a workbook with Aspose.Cells, sets Worksheet.PageSetup.PrintArea to an OFFSET formula that automatically expands as data is added, and saves the file. | Update an existing Aspose.Cells worksheet to apply a COUNTA‑driven OFFSET range as the print area so the printable region updates when new rows or columns are inserted.
// Common Searches: C# Aspose.Cells set print area that expands with new data | How to use OFFSET and COUNTA for dynamic print range in Aspose.Cells | Aspose.Cells PageSetup.PrintArea dynamic range example | Automatically adjust worksheet print area in C# with Aspose.Cells
// Tags: auto expanding print area Aspose.Cells C# | OFFSET formula PageSetup.PrintArea | COUNTA driven print range | worksheet print area automation Aspose.Cells | C# Excel dynamic print range

using Aspose.Cells;
using System;

// Creates a new workbook, defines a dynamic print area using an OFFSET formula that references COUNTA on column A and row 1, assigns it to Worksheet.PageSetup.PrintArea, and saves the workbook as DynamicPrintArea.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // (Optional) Populate some sample data
        // for (int i = 0; i < 10; i++)
        // {
        //     sheet.Cells[i, 0].PutValue($"Item {i + 1}");
        //     sheet.Cells[i, 1].PutValue(i * 10);
        // }

        // Define a dynamic print area that expands with new rows in column A
        // and new columns in row 1 using Excel formulas.
        string dynamicRange = "OFFSET($A$1,0,0,COUNTA($A:$A),COUNTA($1:$1))";

        // Apply the dynamic print area to the worksheet
        sheet.PageSetup.PrintArea = dynamicRange;

        // Save the workbook
        workbook.Save("DynamicPrintArea.xlsx");
    }
}
