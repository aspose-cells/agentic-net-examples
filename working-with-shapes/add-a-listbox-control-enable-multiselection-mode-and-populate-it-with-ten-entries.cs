// Title: Create a multi‑selection ListBox shape in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a ListBox shape on a worksheet, set its InputRange to A1:A10, enable MultiSelect, link the selected items to cell B1, and save the workbook with Aspose.Cells in C#. | Generate an Excel file that writes ten items to A1:A10, adds a ListBox control configured for multiple selections, and writes the chosen values to a linked cell using Aspose.Cells for .NET.
// Common Searches: aspnet cells how to add a ListBox control with multi select to an Excel sheet | c# set ListBox InputRange and linked cell using Aspose.Cells | enable multiple selection for ListBox shape in Aspose.Cells workbook | populate Aspose.Cells ListBox from a range of cells and retrieve selected items
// Tags: Aspose.Cells add ListBox shape C# | ListBox MultiSelect Aspose.Cells | set ListBox InputRange Excel Aspose | link ListBox selected values to cell Aspose.Cells | populate ListBox from cell range C#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, writes ten items into cells A1:A10, adds a ListBox shape at row 2 column 0, sets its InputRange to A1:A10, enables multi‑selection, links the selected items to cell B1, and saves the file as ListBoxExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate ten entries in cells A1:A10
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Add a ListBox control to the worksheet
            // Parameters: upper left row, upper left column, top, left, width, height
            ListBox listBox = sheet.Shapes.AddListBox(2, 0, 0, 0, 150, 120);

            // Set the range that provides the list items
            listBox.InputRange = "A1:A10";

            // Optionally link the selected value(s) to a cell
            listBox.LinkedCell = "B1";

            // Save the workbook
            workbook.Save("ListBoxExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
