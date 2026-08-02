// Title: Aspose.Cells for .NET – Set Print Title Columns A:B in a Worksheet (C#)
// Description: C# example that creates a workbook, accesses the first worksheet, and assigns PageSetup.PrintTitleColumns = "$A:$B" so columns A and B repeat on every printed page. Includes optional sample data and saves the file as PrintTitleColumnsAB.xlsx.
// Keywords: Aspose.Cells | C# | PrintTitleColumns | repeat columns A B | worksheet print titles | PageSetup.PrintTitleColumns | Excel repeat left columns | Aspose.Cells .NET example
// Common Searches: Aspose.Cells repeat columns A B on printed pages | C# set PrintTitleColumns property Aspose.Cells | How to set print title columns in Aspose.Cells | Aspose.Cells PageSetup PrintTitleColumns example | Excel repeat left columns using Aspose.Cells
// Developer Intent: Configure columns A and B to repeat as print titles on each printed page of an Excel worksheet.
// Use Cases: Generate multi‑page reports where the first two columns contain persistent header information. | Programmatically create Excel files with fixed left‑side columns for consistent printing across large datasets. | Demonstrate column repetition by adding sample data before saving the workbook.
// AI Prompts: Show C# code to set both print title rows and columns in an Aspose.Cells worksheet. | Create a method that sets PrintTitleColumns based on a user‑selected column range. | Explain how PrintTitleColumns interacts with PrintArea, FitToPages, and other PageSetup settings.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintTitleColumnsExample
{
    // C# example that creates a workbook, accesses the first worksheet, and assigns PageSetup.PrintTitleColumns = "$A:$B" so columns A and B repeat on every printed page. Includes optional sample data and saves the file as PrintTitleColumnsAB.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set columns A and B to repeat on the left side of each printed page
            worksheet.PageSetup.PrintTitleColumns = "$A:$B";

            // (Optional) Add sample data to demonstrate the repeated columns
            for (int row = 1; row <= 50; row++)
            {
                worksheet.Cells[$"A{row}"].PutValue($"Header A - Row {row}");
                worksheet.Cells[$"B{row}"].PutValue($"Header B - Row {row}");
                worksheet.Cells[$"C{row}"].PutValue($"Data {row}");
            }

            // Save the workbook
            workbook.Save("PrintTitleColumnsAB.xlsx");
        }
    }
}
