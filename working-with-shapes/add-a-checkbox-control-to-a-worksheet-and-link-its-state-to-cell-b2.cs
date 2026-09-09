// Title: Insert a CheckBox form control at B2 and bind its state to cell B2 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that adds a CheckBox shape to the first worksheet at cell B2, links it to cell B2, sets the linked cell to FALSE, and saves the workbook as an .xlsx file with Aspose.Cells. | Create an Aspose.Cells workbook in C# containing a linked CheckBox positioned at B2, initializing the checkbox as unchecked and persisting the file.
// Common Searches: Aspose.Cells C# add a checkbox form control to a specific cell | How to link a checkbox to a cell in an Excel file using Aspose.Cells .NET | Set default state of a linked checkbox to false with Aspose.Cells C# | Save an Excel workbook that includes a checkbox shape using Aspose.Cells for .NET
// Tags: Aspose.Cells add checkbox shape C# | link checkbox to cell Aspose.Cells | initialize linked cell false Aspose.Cells | save workbook with checkbox Aspose.Cells Xlsx | form control checkbox Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program creates a new workbook, inserts a CheckBox form control at cell B2, links the CheckBox to cell B2, initializes the linked cell to FALSE (unchecked), and saves the file as CheckBoxExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a CheckBox control at cell B2 (row index 1, column index 1)
        // Parameters: upperLeftRow, upperLeftColumn, top, left, width, height
        CheckBox checkBox = sheet.Shapes.AddCheckBox(1, 1, 0, 0, 100, 20);

        // Link the CheckBox state to cell B2
        checkBox.LinkedCell = "B2";

        // Set the initial value of B2 (FALSE = unchecked, TRUE = checked)
        sheet.Cells["B2"].PutValue(false);

        // Save the workbook
        workbook.Save("CheckBoxExample.xlsx", SaveFormat.Xlsx);
    }
}
