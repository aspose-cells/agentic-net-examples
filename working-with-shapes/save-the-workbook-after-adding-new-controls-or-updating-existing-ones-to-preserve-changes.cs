// Title: Add, Update, and Save a CheckBox ActiveX Control in Excel with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, insert a CheckBox ActiveX control, set its properties, save the file, reload it, modify the control's caption, and save the changes again using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# ActiveX control | add CheckBox ActiveX Excel | update ActiveX caption | save workbook after modifying shapes | .NET Excel ActiveX example | preserve ActiveX controls Aspose
// Common Searches: How to add a CheckBox ActiveX control with Aspose.Cells C# | Update ActiveX control properties and save workbook in .NET | Saving changes to Excel ActiveX controls using Aspose.Cells | Load Excel file, modify ActiveX control, and resave with C# | Aspose.Cells example for editing shapes
// Developer Intent: Insert a CheckBox ActiveX control into a new workbook, change its caption later, and ensure all modifications are persisted when the file is saved.
// Use Cases: Generate Excel templates that include interactive CheckBox controls for user input. | Programmatically adjust control captions based on dynamic business data before distribution. | Automate report creation with embedded ActiveX controls that remain functional after saving.
// AI Prompts: Write C# code with Aspose.Cells to add a RadioButton ActiveX control, set its properties, and save the workbook. | Provide a step‑by‑step guide to locate a specific ActiveX control by name in an existing Excel file, change its value, and save the changes using Aspose.Cells. | Explain how to iterate over all ActiveX controls on a worksheet and modify a common attribute (e.g., Enabled) with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsControlSaveDemo
{
    // Demonstrates how to create a workbook, insert a CheckBox ActiveX control, set its properties, save the file, reload it, modify the control's caption, and save the changes again using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add an ActiveX control
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create a new workbook
            Worksheet sheet = workbook.Worksheets[0];               // get the first worksheet

            // Add a CheckBox ActiveX control to the worksheet
            Shape shape = sheet.Shapes.AddActiveXControl(
                ControlType.CheckBox,   // type of control
                1, 0, 1, 0,             // upper left row, column, lower right row, column
                120, 30);               // width and height in points

            // Cast to the specific control type to set its properties
            CheckBoxActiveXControl checkBox = (CheckBoxActiveXControl)shape.ActiveXControl;
            checkBox.Caption = "Demo CheckBox";
            checkBox.IsAutoSize = true;

            // Save the workbook to preserve the newly added control
            workbook.Save("WorkbookWithControl.xlsx"); // uses Save(string) overload

            // -------------------------------------------------
            // 2. Load the saved workbook, modify the control, and save again
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook("WorkbookWithControl.xlsx"); // load existing file
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

            // Retrieve the first shape (our CheckBox) and change its caption
            Shape loadedShape = loadedSheet.Shapes[0];
            if (loadedShape.ActiveXControl is CheckBoxActiveXControl loadedCheckBox)
            {
                loadedCheckBox.Caption = "Updated CheckBox Caption";
            }

            // Save the changes back to disk
            loadedWorkbook.Save("WorkbookWithControl_Updated.xlsx"); // preserve updates
        }
    }
}
