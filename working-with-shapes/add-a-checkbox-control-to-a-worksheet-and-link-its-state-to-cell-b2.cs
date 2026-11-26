// Title: Create a CheckBox form control in an Excel worksheet and bind it to cell B2 using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add checkbox to cell B2 | How to link a checkbox control to a worksheet cell with Aspose.Cells .NET | Set initial checked state for a checkbox shape in Aspose.Cells workbook | Programmatically create Excel form controls using Aspose.Cells C# | Save Excel file with linked checkbox using Aspose.Cells SaveFormat Xlsx
// Tags: Aspose.Cells add checkbox shape | checkbox linked cell Aspose.Cells | C# Aspose.Cells form control binding | initialize checkbox value true Aspose.Cells | save workbook with form controls Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example demonstrates how to create a new workbook, insert a CheckBox form control positioned at cell B2 on the first worksheet, bind the CheckBox to cell B2 so its TRUE/FALSE state is stored there, set the CheckBox to be initially checked, and save the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define the position and size of the checkbox.
            // Row and column are zero‑based. B2 corresponds to row 1, column 1.
            int upperLeftRow = 1;      // B2 row
            int upperLeftColumn = 1;   // B2 column
            int height = 20;           // Height in points
            int width = 100;           // Width in points

            // Add the checkbox to the worksheet.
            // The Add method returns the index of the newly added checkbox.
            int checkBoxIndex = sheet.CheckBoxes.Add(upperLeftRow, upperLeftColumn, height, width);
            CheckBox checkBox = sheet.CheckBoxes[checkBoxIndex];

            // Link the checkbox state to cell B2.
            // When the checkbox is checked/unchecked, the linked cell will contain TRUE/FALSE.
            checkBox.LinkedCell = "B2";

            // Optionally set an initial state (checked).
            checkBox.Value = true;

            // Save the workbook.
            string outputFile = "CheckBoxLinkedToB2.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
