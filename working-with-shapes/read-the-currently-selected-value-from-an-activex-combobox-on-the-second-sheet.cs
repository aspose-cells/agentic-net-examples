// Title: Read the selected value of an ActiveX ComboBox on the second worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# read selected item from ActiveX ComboBox on sheet index 1 | Get value of an ActiveX ComboBox control in the second worksheet with Aspose.Cells | Retrieve linked cell content of a ComboBoxActiveXControl using Aspose.Cells in C# | Iterate worksheet shapes to locate ComboBoxActiveXControl with Aspose.Cells .NET | Read ActiveX ComboBox selected value from an Excel file via Aspose.Cells library
// Tags: Aspose.Cells read ActiveX ComboBox value C# | ComboBoxActiveXControl selected value extraction | second worksheet shape enumeration Aspose.Cells | linked cell retrieval from ComboBoxActiveXControl | C# console output Excel ActiveX control data

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace ActiveXComboBoxReader
{
    // The example loads an Excel workbook, accesses the second worksheet, scans its Shapes collection to find the first ComboBoxActiveXControl, reads the control's Value and LinkedCell properties, optionally reads the linked cell's content, and prints the worksheet name, linked cell address, linked cell value, and the ComboBox's selected value to the console.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook that contains the ActiveX ComboBox.
            // Adjust the file name/path as needed.
            const string inputFile = "SampleWorkbook.xlsx";

            // Load the workbook.
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook \"{inputFile}\": {ex.Message}");
                return;
            }

            // Ensure the workbook has at least two worksheets.
            if (workbook.Worksheets.Count < 2)
            {
                Console.WriteLine("The workbook does not contain a second worksheet.");
                return;
            }

            // Get the second worksheet (index 1, zero‑based).
            Worksheet secondSheet = workbook.Worksheets[1];

            // Find the first ActiveX ComboBox control on the sheet.
            ComboBoxActiveXControl comboBoxControl = null;
            foreach (Shape shape in secondSheet.Shapes)
            {
                // The Shape.ActiveXControl property returns an ActiveXControl instance
                // if the shape hosts an ActiveX control; otherwise it is null.
                if (shape.ActiveXControl != null && shape.ActiveXControl is ComboBoxActiveXControl)
                {
                    comboBoxControl = (ComboBoxActiveXControl)shape.ActiveXControl;
                    break; // Stop after the first ComboBox is found.
                }
            }

            if (comboBoxControl == null)
            {
                Console.WriteLine("No ActiveX ComboBox control was found on the second worksheet.");
                return;
            }

            // The Value property of ComboBoxActiveXControl holds the currently selected item.
            // It may be null if nothing is selected.
            object selectedValue = comboBoxControl.Value;

            // If the control is linked to a cell, the linked cell also reflects the selection.
            // We'll display both for completeness.
            string linkedCellAddress = comboBoxControl.LinkedCell;
            string linkedCellValue = string.Empty;
            if (!string.IsNullOrEmpty(linkedCellAddress))
            {
                try
                {
                    Cell linkedCell = secondSheet.Cells[linkedCellAddress];
                    linkedCellValue = linkedCell.Value?.ToString() ?? string.Empty;
                }
                catch
                {
                    // Ignore errors reading the linked cell.
                }
            }

            Console.WriteLine("=== ActiveX ComboBox Information ===");
            Console.WriteLine($"Worksheet: {secondSheet.Name}");
            Console.WriteLine($"ComboBox Linked Cell: {linkedCellAddress}");
            Console.WriteLine($"Linked Cell Value   : {linkedCellValue}");
            Console.WriteLine($"Selected Value      : {selectedValue ?? "(none)"}");
        }
    }
}
