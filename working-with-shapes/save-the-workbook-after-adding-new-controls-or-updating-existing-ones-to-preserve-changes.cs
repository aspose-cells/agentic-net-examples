// Title: Add an ActiveX CheckBox to an Excel worksheet and save the workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, insert an ActiveX CheckBox control with a custom caption and unchecked state, write text to a cell, ensure the output folder exists, and save the file as an .xlsx document while preserving the control.
// Keywords: Aspose.Cells | ActiveX CheckBox | C# | .NET | AddActiveXControl | CheckBoxActiveXControl | save workbook | Excel worksheet | directory creation | persist ActiveX control
// Common Searches: Aspose.Cells add ActiveX CheckBox C# | Save workbook after inserting ActiveX controls Aspose.Cells | Create Excel form with ActiveX controls using .NET | Ensure output folder exists before saving Excel file C# | How to set checkbox value with Aspose.Cells
// Developer Intent: Insert an ActiveX CheckBox into a worksheet and persist the workbook with the new control.
// Use Cases: Generate a terms‑acceptance checkbox on a programmatically created report and distribute the file. | Build a form‑style worksheet containing multiple ActiveX controls (e.g., CheckBox, ComboBox) and save it as a reusable template. | Update cell content based on a control’s state before saving the workbook for downstream processing.
// AI Prompts: Write C# code that adds several ActiveX controls (CheckBox, ComboBox, ListBox) to a worksheet, configures their properties, and saves the workbook with Aspose.Cells. | Show how to load an existing .xlsx file, modify the caption of an existing ActiveX CheckBox, and save the changes using Aspose.Cells for .NET. | Explain best practices for directory creation, error handling, and workbook saving when the file contains ActiveX controls.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsControlSaveDemo
{
    // Demonstrates how to create a new Workbook, insert an ActiveX CheckBox control with a custom caption and unchecked state, write text to a cell, ensure the output folder exists, and save the file as an .xlsx document while preserving the control.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                using (Workbook workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet sheet = workbook.Worksheets[0];

                    // Add an ActiveX CheckBox control to the worksheet
                    // Parameters: control type, upper left row, upper left column, lower right row, lower right column, width, height
                    Shape shape = sheet.Shapes.AddActiveXControl(
                        ControlType.CheckBox, 1, 0, 1, 0, 100, 30);
                    CheckBoxActiveXControl checkBox = (CheckBoxActiveXControl)shape.ActiveXControl;

                    // Set properties of the CheckBox control
                    checkBox.Caption = "Accept Terms";

                    // Set the initial state (unchecked). Use numeric cast to avoid enum member mismatch.
                    checkBox.Value = (CheckValueType)0; // 0 = Unchecked

                    // Update a cell to reflect the control state (example of updating existing content)
                    sheet.Cells["A1"].PutValue("User must accept terms before proceeding.");

                    // Define output path
                    string outputPath = "ControlDemo.xlsx";

                    // Ensure the directory for the output file exists
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // Save the workbook with the added control
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully at '{outputPath}' with ActiveX control.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
