// Title: Set Worksheet Print Area to a Named Range (ReportArea) with Aspose.Cells for .NET
// Description: Shows how to create a named range called ReportArea, assign it to a worksheet’s PageSetup.PrintArea, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | PrintArea | named range | PageSetup | dynamic printing | worksheet print area | set print area programmatically
// Common Searches: Aspose.Cells set print area named range C# | how to assign named range to PrintArea in Aspose.Cells | dynamic worksheet print area Aspose.Cells .NET | PageSetup.PrintArea example with named range | C# Aspose.Cells print area from named range
// Developer Intent: Define a named range and use it as the worksheet’s print area for flexible, programmatic printing.
// Use Cases: Generate reports where the printable region expands automatically by updating the named range before printing. | Create invoices that print only the area containing invoice details by assigning a named range as the print area. | Automate dashboard printing by setting a specific named range as the print area for each worksheet.
// AI Prompts: Adjust the code to calculate the named range size from the last used row before setting the PrintArea. | Provide an example that assigns different named ranges as print areas for multiple worksheets in one workbook. | Explain how to read the current PrintArea name at runtime and switch it to another named range.

using System;
using Aspose.Cells;

// Shows how to create a named range called ReportArea, assign it to a worksheet’s PageSetup.PrintArea, and save the workbook using Aspose.Cells for .NET.
class SetPrintAreaToNamedRange
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Fill the worksheet with sample data (optional)
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 5; col++)
            {
                sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
            }
        }

        // Define the cell area that should be printed
        string printArea = "A1:E20";

        // Create a named range called "ReportArea" that refers to the defined area
        int nameIdx = workbook.Worksheets.Names.Add("ReportArea");
        workbook.Worksheets.Names[nameIdx].RefersTo = $"={sheet.Name}!{printArea}";

        // Assign the named range as the worksheet's print area
        sheet.PageSetup.PrintArea = "ReportArea";

        // Save the workbook
        workbook.Save("ReportAreaDemo.xlsx");
    }
}
