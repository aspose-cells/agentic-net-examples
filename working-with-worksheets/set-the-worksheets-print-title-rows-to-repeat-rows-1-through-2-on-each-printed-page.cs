// Title: Set rows 1‑2 as repeating print titles on an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to configure the PageSetup.PrintTitleRows property so that rows 1 and 2 repeat on every printed page. | Write C# code that loads a workbook, sets the print title rows to $1:$2, and saves the file with Aspose.Cells. | Demonstrate how to apply an A1‑style range to the print title rows setting for a worksheet in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set PrintTitleRows to repeat first two rows on each page | how to make header rows repeat when printing Excel with Aspose.Cells .NET | C# example for worksheet page setup repeat rows using Aspose.Cells | specify A1 notation for print title rows in Aspose.Cells workbook | repeat rows 1 and 2 on printed pages Aspose.Cells page setup C#
// Tags: Aspose.Cells page setup print titles | C# worksheet repeat rows on print | Excel print title rows with Aspose.Cells | A1 style range for print titles Aspose.Cells | save workbook with repeated header rows C#

using Aspose.Cells;

// The program creates a new workbook, accesses the first worksheet, sets rows 1‑2 as print title rows using the PageSetup.PrintTitleRows property (A1 notation "$1:$2"), and saves the workbook as PrintTitleRowsExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // creates a new workbook with one default worksheet

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the print title rows to repeat rows 1 through 2 on each printed page
        // The range is specified in A1 style notation
        worksheet.PageSetup.PrintTitleRows = "$1:$2";

        // Save the workbook to a file
        workbook.Save("PrintTitleRowsExample.xlsx");
    }
}
