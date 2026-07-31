// Title: Set Print Area A1:M50 with Aspose.Cells for .NET (C#)
// Description: This C# example creates a new Workbook, accesses the first Worksheet, assigns the PageSetup.PrintArea property to "A1:M50" to restrict the printable range, and saves the file as PrintAreaDemo.xlsx.
// Keywords: Aspose.Cells C# print area | set print area A1:M50 | PageSetup.PrintArea example | limit printed range Aspose.Cells | save workbook with print area | Aspose.Cells .NET printing
// Common Searches: Aspose.Cells set print area .NET | C# define print area A1:M50 | how to limit printable range in Aspose.Cells | PageSetup.PrintArea usage C# | print specific cells Aspose.Cells
// Developer Intent: Define a specific printable region (A1:M50) for a worksheet using Aspose.Cells in C#.
// Use Cases: Generate a sales report where only cells A1:M50 should appear on printed pages. | Create invoices that print only the populated area without extra blank pages. | Prepare a template with a fixed printable area to ensure consistent page layout across users. | Automate batch printing of worksheets with a predefined print range.
// AI Prompts: Show C# code to apply the same print area to all worksheets in a workbook using Aspose.Cells. | How can I read the current PrintArea of a worksheet and change it to A1:M50 with Aspose.Cells? | Provide an example that sets different print areas for multiple worksheets in one workbook (C# Aspose.Cells). | Explain how to reset or clear the print area after it has been set in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaDemo
{
    // This C# example creates a new Workbook, accesses the first Worksheet, assigns the PageSetup.PrintArea property to "A1:M50" to restrict the printable range, and saves the file as PrintAreaDemo.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Define the print area from A1 to M50
            worksheet.PageSetup.PrintArea = "A1:M50";

            // Save the workbook (lifecycle rule: save)
            workbook.Save("PrintAreaDemo.xlsx");

            Console.WriteLine("Workbook saved with print area A1:M50.");
        }
    }
}
