// Title: C# – Add an ActiveX ComboBox to an Aspose.Cells worksheet, bind it to a range, and set the drop‑down width
// Description: Demonstrates how to create a new workbook with Aspose.Cells, write five items to column A, insert a ComboBox ActiveX control, link it to the range A1:A5 using ListFillRange, configure the drop‑down width with ListWidth, set the visible row count via ListRows, and save the file as an XLSX document.
// Keywords: Aspose.Cells ComboBox ActiveX | C# Aspose.Cells ComboBox | .NET add ComboBox to worksheet | ListFillRange property | ListWidth drop‑down width | ListRows visible rows | populate ComboBox from cells | ActiveX control Excel C# | Aspose.Cells shape AddActiveXControl
// Common Searches: Aspose.Cells set ComboBox drop‑down width C# | How to bind ComboBox to cell range in Aspose.Cells | Add ActiveX ComboBox to Excel worksheet using Aspose.Cells | ListWidth and ListRows properties for ComboBox in Aspose.Cells | Create ComboBox with items from column A in C#
// Developer Intent: Insert an ActiveX ComboBox into a worksheet, connect it to a cell range, and customize its drop‑down width and visible rows using Aspose.Cells for .NET.
// Use Cases: Design a data‑entry template where users pick options from a ComboBox whose list is wide enough for long text. | Generate reports that include a ComboBox linked to a dynamic range, ensuring consistent UI across different locales. | Automate spreadsheet creation with predefined selection lists, controlling the number of visible items and drop‑down width for better readability.
// AI Prompts: Write C# code with Aspose.Cells to add a ComboBox ActiveX control, bind it to cells A1:A5, set ListWidth to 180 points, and save the workbook. | Explain the effect of ListFillRange, ListWidth, and ListRows on an ActiveX ComboBox in an Aspose.Cells worksheet. | Provide a step‑by‑step tutorial for creating a ComboBox with a custom drop‑down width and a specific number of visible rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsComboBoxDemo
{
    // Demonstrates how to create a new workbook with Aspose.Cells, write five items to column A, insert a ComboBox ActiveX control, link it to the range A1:A5 using ListFillRange, configure the drop‑down width with ListWidth, set the visible row count via ListRows, and save the file as an XLSX document.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate five items in column A (the list for the combo box)
            worksheet.Cells["A1"].PutValue("Item 1");
            worksheet.Cells["A2"].PutValue("Item 2");
            worksheet.Cells["A3"].PutValue("Item 3");
            worksheet.Cells["A4"].PutValue("Item 4");
            worksheet.Cells["A5"].PutValue("Item 5");

            // Add a ComboBox ActiveX control to the worksheet
            // Parameters: ControlType, upper left row, upper left column, top offset, left offset, height, width
            Shape shape = worksheet.Shapes.AddActiveXControl(
                ControlType.ComboBox, // control type
                2,    // upper left row index (zero‑based)
                0,    // upper left column index
                0,    // top offset in pixels
                0,    // left offset in pixels
                30,   // height in pixels
                100   // width in pixels
            );

            // Cast the ActiveXControl to ComboBoxActiveXControl to access its specific properties
            ComboBoxActiveXControl comboBox = (ComboBoxActiveXControl)shape.ActiveXControl;

            // Link the combo box to the range containing the items
            comboBox.ListFillRange = "A1:A5";

            // Define the drop‑down width (in points)
            comboBox.ListWidth = 150; // Adjust as needed

            // Optionally set how many rows are visible when the list drops down
            comboBox.ListRows = 5;

            // Save the workbook to a file
            workbook.Save("ComboBoxWithDropDownWidth.xlsx");
        }
    }
}
