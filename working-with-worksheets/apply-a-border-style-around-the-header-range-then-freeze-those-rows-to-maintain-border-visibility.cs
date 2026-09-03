// Title: Apply a thick blue outline border to a header range and freeze the header row in an Excel file using Aspose.Cells for .NET
// AI Prompts: Create a header range A1:C1, set a thick blue outline border, and freeze the first row with FreezePanes in Aspose.Cells C#. | Generate an Excel workbook that styles the header row with a blue thick border and keeps it visible while scrolling using Aspose.Cells for .NET. | Write C# code to apply outline borders to a range and freeze panes so the header stays fixed in an Aspose.Cells spreadsheet.
// Common Searches: how to add a thick blue outline border to a header row using Aspose.Cells C# | freeze first row after styling header in Aspose.Cells .NET | Aspose.Cells C# set outline borders for a range and freeze panes together
// Tags: setoutlineborders Aspose.Cells range | freezepanes header row Aspose.Cells | apply thick blue border Aspose.Cells | styled header Excel Aspose.Cells C# | save workbook with frozen header Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsHeaderBorderFreeze
{
    // // This program creates a new workbook, writes header values to A1:C1, applies a thick blue outline border to that range, freezes the first row, and saves the file as HeaderBorderFreeze.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample header data (first row)
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["C1"].PutValue("Header3");

                // Define the header range (first row, columns A to C)
                AsposeRange headerRange = cells.CreateRange("A1:C1");

                // Apply a thick blue outline border around the header range
                headerRange.SetOutlineBorders(CellBorderType.Thick, Color.Blue);

                // Freeze the header row so the border remains visible while scrolling
                // Freeze at row index 1 (the row after the header), column index 0,
                // with 1 frozen row and 0 frozen columns.
                worksheet.FreezePanes(1, 0, 1, 0);

                // Save the workbook
                workbook.Save("HeaderBorderFreeze.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
