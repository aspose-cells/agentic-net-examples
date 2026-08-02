// Title: Aspose.Cells for .NET – Assign Print Area from a Named Range (ReportArea)
// Description: Shows how to create a workbook, define a named range called ReportArea (A1:C10), and set the worksheet’s PageSetup.PrintArea to that range so only the specified cells are printed or exported.
// Keywords: Aspose.Cells print area | named range print area C# | PageSetup.PrintArea | define named range Aspose.Cells | dynamic print area .NET | set print area Aspose.Cells | ReportArea named range
// Common Searches: Aspose.Cells set print area named range | C# Aspose.Cells print specific cells | PageSetup.PrintArea using named range | define ReportArea in Aspose.Cells | dynamic print area Aspose.Cells .NET
// Developer Intent: Set the worksheet’s print area to the predefined named range ReportArea.
// Use Cases: Generate a report workbook and limit printing to the data block defined by ReportArea. | Apply the same named range as the print area across multiple sheets for consistent layout. | Modify the cells in ReportArea and have the print area automatically reflect the changes without extra code.
// AI Prompts: Write C# code that creates a named range and assigns it to PageSetup.PrintArea with Aspose.Cells. | Explain how updating an existing named range automatically updates the worksheet’s print area in Aspose.Cells. | Provide a step‑by‑step tutorial for defining a named range and using it as the print area in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    // Shows how to create a workbook, define a named range called ReportArea (A1:C10), and set the worksheet’s PageSetup.PrintArea to that range so only the specified cells are printed or exported.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data (A1:C10)
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define a named range called "ReportArea" that refers to A1:C10
            int nameIndex = workbook.Worksheets.Names.Add("ReportArea");
            // The RefersTo string must start with '=' and include the sheet name
            workbook.Worksheets.Names[nameIndex].RefersTo = $"=Sheet1!$A$1:$C$10";

            // Set the worksheet's print area to the named range
            // PrintArea can accept the name of a defined range
            sheet.PageSetup.PrintArea = "ReportArea";

            // Save the workbook (the print area will be used when printing or exporting)
            workbook.Save("ReportAreaDemo.xlsx");

            Console.WriteLine("Workbook created with print area set to named range 'ReportArea'.");
        }
    }
}
