// Title: Unlock TextBox and ActiveX TextBox Shapes in Excel with Aspose.Cells for .NET
// Description: Learn how to programmatically unlock regular TextBox shapes and ActiveX TextBox controls in an Excel workbook, enable object editing, and save the changes using Aspose.Cells C#.
// Keywords: Aspose.Cells C# unlock TextBox shape | ActiveX TextBox unlock Aspose.Cells | worksheet protection AllowEditingObject | shape.IsLocked property | modify Excel textbox position programmatically | unlock drawing objects Aspose.Cells | C# Excel shape manipulation
// Common Searches: C# unlock locked TextBox shape in Excel using Aspose.Cells | How to enable moving and resizing of TextBox in protected worksheet Aspose.Cells | Unlock ActiveX TextBox control programmatically with Aspose.Cells .NET | Set shape.IsLocked false for all textboxes in a workbook | Allow editing of drawing objects in protected Excel sheet Aspose.Cells
// Developer Intent: Programmatically remove lock restrictions from TextBox and ActiveX TextBox shapes so their position, size, and content can be edited in an Excel file.
// Use Cases: Prepare a template workbook by unlocking textboxes before populating data. | Allow end‑users to edit ActiveX TextBox inputs on a protected sheet after distribution. | Automate layout adjustments by batch‑unlocking all textbox shapes across multiple workbooks. | Integrate into a reporting pipeline to reposition textboxes without manual intervention. | Create a utility that toggles shape lock status based on worksheet protection settings.
// AI Prompts: Generate C# code with Aspose.Cells that iterates through all shapes, unlocks TextBox and ActiveX TextBox controls, sets worksheet.Protection.AllowEditingObject = true, and saves the workbook. | Explain how shape.IsLocked and worksheet.Protection.AllowEditingObject interact when unlocking drawing objects in Aspose.Cells. | Write a reusable method UnlockTextBoxes(string inputPath, string outputPath) that only unlocks regular TextBox shapes, leaving other shapes unchanged. | Provide a step‑by‑step guide to batch‑unlock textbox shapes in multiple Excel files using Aspose.Cells and parallel processing. | Show how to check if a shape is a TextBoxActiveXControl before unlocking it in C#.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

// Learn how to programmatically unlock regular TextBox shapes and ActiveX TextBox controls in an Excel workbook, enable object editing, and save the changes using Aspose.Cells C#.
class UnlockTextBox
{
    static void Main()
    {
        // Load the workbook that contains the locked textbox
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Iterate through all shapes on the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // If the shape is a regular TextBox, unlock it so its position/size can be edited
            if (shape is TextBox)
            {
                shape.IsLocked = false;
            }

            // If the shape hosts an ActiveX TextBox control, unlock the control and make it editable
            if (shape.ActiveXControl is TextBoxActiveXControl activeX)
            {
                activeX.IsLocked = false;      // Unlock data editing
                activeX.IsEditable = true;     // Allow typing into the control
            }
        }

        // Ensure the worksheet protection permits manipulation of drawing objects
        worksheet.Protection.AllowEditingObject = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
