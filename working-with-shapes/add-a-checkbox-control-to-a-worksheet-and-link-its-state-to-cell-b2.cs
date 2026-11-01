// Title: Create a CheckBox form control in an Excel worksheet and link it to cell B2 using Aspose.Cells for .NET (C#)
// Common Searches: asp.net add checkbox shape to Excel worksheet with Aspose.Cells C# | how to bind a checkbox control to a specific cell in an Excel file using Aspose.Cells | programmatically insert a form control checkbox and set linked cell B2 in C# | Aspose.Cells example linking CheckBox state to worksheet cell | C# code to add a CheckBox shape and associate it with cell B2 in .xlsx
// Tags: Aspose.Cells AddCheckBox shape C# | checkbox linked cell B2 Aspose.Cells | Excel form control insertion .NET | bind checkbox state to worksheet cell | create checkbox shape in .xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // The sample creates a new Workbook, adds a CheckBox shape to the first worksheet, sets its display text, links the CheckBox's state to cell B2, and saves the file as CheckBoxDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                var workbook = new Workbook();

                // Get the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Add a CheckBox shape: upper left row, upper left column, top offset, left offset, width, height
                var checkBox = worksheet.Shapes.AddCheckBox(0, 0, 0, 0, 100, 20);

                // Set the text displayed next to the CheckBox (optional)
                checkBox.Text = "Check me";

                // Link the CheckBox state to cell B2
                checkBox.LinkedCell = "B2";

                // Define output file path
                string outputPath = "CheckBoxDemo.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
