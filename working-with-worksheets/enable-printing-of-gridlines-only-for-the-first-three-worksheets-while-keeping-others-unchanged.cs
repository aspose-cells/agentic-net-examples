// Title: Aspose.Cells C# – Print Gridlines on the First Three Worksheets Only
// Description: Creates a workbook, adds worksheets to ensure at least five sheets, then enables PageSetup.PrintGridlines for the first three worksheets while leaving the remaining sheets unchanged, and saves the file.
// Keywords: Aspose.Cells PrintGridlines C# | enable gridlines printing specific worksheets | PageSetup.PrintGridlines property | Excel gridlines first three sheets | Aspose.Cells multiple worksheets settings
// Common Searches: Aspose.Cells print gridlines on selected sheets | C# set PrintGridlines for first three worksheets | how to enable gridlines printing for some worksheets Aspose.Cells | Aspose.Cells PageSetup.PrintGridlines example
// Developer Intent: Set the PrintGridlines flag only on the first three worksheets of an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Financial reports where summary and totals sheets need printed gridlines, but detailed data sheets stay clean. | Templates that show gridlines on overview pages while departmental sheets are printed without them. | Mixed‑audience workbooks that require gridlines on introductory worksheets and no gridlines on the rest.
// AI Prompts: Generate C# code to enable PrintGridlines on the first N worksheets in Aspose.Cells. | Show how to toggle PageSetup.PrintGridlines for selected worksheets while leaving others unchanged. | Explain the impact of the PrintGridlines setting on worksheet printing in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a workbook, adds worksheets to ensure at least five sheets, then enables PageSetup.PrintGridlines for the first three worksheets while leaving the remaining sheets unchanged, and saves the file.
class EnableGridlinesFirstThreeSheets
{
    static void Main()
    {
        // Create a new workbook (default has one worksheet)
        Workbook workbook = new Workbook();

        // Add additional worksheets so we have at least five sheets
        while (workbook.Worksheets.Count < 5)
        {
            workbook.Worksheets.Add();
        }

        // Enable printing of gridlines for the first three worksheets only
        for (int i = 0; i < Math.Min(3, workbook.Worksheets.Count); i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            sheet.PageSetup.PrintGridlines = true; // gridlines will be printed when this sheet is printed
        }

        // Save the workbook to a file
        workbook.Save("GridlinesFirstThreeSheets.xlsx");
    }
}
