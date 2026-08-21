// Title: C# – Add a CheckBox to an Excel worksheet and link it to cell B2 with Aspose.Cells
// Description: Demonstrates how to create a new Workbook, insert a CheckBox shape at row 1 column 1, set its caption, bind the checkbox state to cell B2 using the LinkedCell property, define the initial checked value, and save the file as CheckboxLinkedToB2.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells checkbox example | link checkbox to cell | Aspose.Cells C# CheckBox | LinkedCell property | Excel form controls Aspose
// Common Searches: add checkbox to Excel with Aspose.Cells C# | link checkbox state to a cell using Aspose | set initial value of Aspose.Cells checkbox | read linked cell value from Aspose.Cells checkbox | Aspose.Cells sample code for form controls
// Developer Intent: Insert a CheckBox shape into a worksheet and bind its checked state to cell B2 programmatically.
// Use Cases: Create interactive Excel forms where a user’s acceptance updates a flag in B2 for downstream calculations. | Generate templates that toggle optional sections based on the linked checkbox value in B2. | Automate report generation where the presence of a checkbox controls conditional formatting or data inclusion.
// AI Prompts: Write C# code with Aspose.Cells to place a checkbox at D5 and link it to cell C10. | Show how to retrieve the value of a checkbox's linked cell after opening the workbook with Aspose.Cells. | Explain how to change the checkbox caption and default checked state using Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a new Workbook, insert a CheckBox shape at row 1 column 1, set its caption, bind the checkbox state to cell B2 using the LinkedCell property, define the initial checked value, and save the file as CheckboxLinkedToB2.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a checkbox at row 1, column 1 (zero‑based indexes)
        // Height = 20 pixels, Width = 100 pixels
        int checkBoxIndex = sheet.CheckBoxes.Add(1, 1, 20, 100);
        CheckBox checkBox = sheet.CheckBoxes[checkBoxIndex];

        // Optional: set the displayed text of the checkbox
        checkBox.Text = "Accept";

        // Link the checkbox state to cell B2
        checkBox.LinkedCell = "B2";

        // Optional: set the initial checked state
        checkBox.Value = false;

        // Save the workbook
        workbook.Save("CheckboxLinkedToB2.xlsx");
    }
}
