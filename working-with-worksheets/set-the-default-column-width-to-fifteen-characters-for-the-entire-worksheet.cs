// Title: C# – Set default column width to 15 characters for an entire worksheet with Aspose.Cells
// Description: This .NET example creates a new Workbook, accesses the first Worksheet, and sets the worksheet's default column width to 15 characters by assigning 15.0 to the Cells.StandardWidth property. The code prints the applied width, verifies the first column's actual width, and saves the file as "DefaultColumnWidth.xlsx".
// Keywords: Aspose.Cells default column width | Cells.StandardWidth C# | set worksheet column width characters | Aspose.Cells column width example | C# spreadsheet default width
// Common Searches: Aspose.Cells set default column width .NET | How to use Cells.StandardWidth in C# | Set column width for all columns Aspose.Cells | Default column width in characters Aspose.Cells
// Developer Intent: Set the worksheet’s default column width to 15 characters using Aspose.Cells for .NET.
// Use Cases: Create a template where every column starts with a uniform width before applying custom adjustments. | Generate reports that require consistent column sizing for readability and printing. | Prepare spreadsheets for export to other systems that expect a predefined column width.
// AI Prompts: Show how to change the default column width to 20 characters and then set column B to 30 characters using Aspose.Cells for .NET. | Write code that reads the current StandardWidth from a workbook, asks the user for a new width, and updates the worksheet accordingly. | Explain the relationship between the StandardWidth property and pixel units, including how to convert character width to pixels for different DPI settings.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This .NET example creates a new Workbook, accesses the first Worksheet, and sets the worksheet's default column width to 15 characters by assigning 15.0 to the Cells.StandardWidth property. The code prints the applied width, verifies the first column's actual width, and saves the file as "DefaultColumnWidth.xlsx".
    public class SetDefaultColumnWidth
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the default column width for the entire worksheet to 15 characters
            // This uses the Cells.StandardWidth property which defines the default width in character units
            worksheet.Cells.StandardWidth = 15.0;

            // Optional: verify the setting
            Console.WriteLine("Standard Width set to: " + worksheet.Cells.StandardWidth);
            Console.WriteLine("Column 0 actual width: " + worksheet.Cells.GetColumnWidth(0));

            // Save the workbook (lifecycle: save)
            workbook.Save("DefaultColumnWidth.xlsx");
        }
    }
}
