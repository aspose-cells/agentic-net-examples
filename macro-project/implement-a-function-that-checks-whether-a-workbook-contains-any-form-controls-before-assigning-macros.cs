// Title: Detect Form Controls (CheckBoxes & ActiveX) in an Aspose.Cells Workbook Before Enabling Macros (C#)
// Description: C# example that scans every worksheet in an Aspose.Cells workbook for legacy CheckBoxes and ActiveX shapes. The ContainsFormControls method returns true when any form control is found, allowing you to conditionally enable macros or apply security policies.
// Keywords: Aspose.Cells form control detection | C# check for checkboxes in workbook | ActiveX shape detection Aspose.Cells | prevent macro assignment Excel | Workbook.ContainsFormControls | EnableMacros conditional | scan worksheets for form controls
// Common Searches: how to detect checkboxes in Aspose.Cells workbook | Aspose.Cells find ActiveX controls before enabling macros | C# iterate worksheets to check for form controls | prevent macro injection when Excel file has form controls | Aspose.Cells check for legacy form controls
// Developer Intent: Identify whether an Excel workbook contains any form controls so that macros can be enabled only when the file is free of interactive elements.
// Use Cases: Validate incoming Excel files and skip macro injection if checkboxes or ActiveX controls are present. | Automate batch processing that enables macros only for clean workbooks. | Create a pre‑publish audit that flags worksheets containing form controls for manual review.
// AI Prompts: Generate a method that returns the names of worksheets containing any form controls (checkboxes, ActiveX) using Aspose.Cells. | Extend ContainsFormControls to also detect combo boxes, list boxes, and option buttons in both legacy and ActiveX collections. | Write unit tests for ContainsFormControls covering scenarios with no controls, only checkboxes, only ActiveX shapes, and a mix of both.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# example that scans every worksheet in an Aspose.Cells workbook for legacy CheckBoxes and ActiveX shapes. The ContainsFormControls method returns true when any form control is found, allowing you to conditionally enable macros or apply security policies.
public class FormControlChecker
{
    // Returns true if the workbook contains any form controls (checkboxes or ActiveX controls)
    public static bool ContainsFormControls(Workbook workbook)
    {
        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check for legacy form controls (CheckBox collection)
            if (sheet.CheckBoxes.Count > 0)
                return true;

            // Check for ActiveX controls added as shapes
            foreach (Shape shape in sheet.Shapes)
            {
                // If the shape has an associated ActiveX control, a form control exists
                if (shape.ActiveXControl != null)
                    return true;
            }
        }

        // No form controls found
        return false;
    }

    // Example method demonstrating the check before enabling macros
    public static void RunExample()
    {
        // Create a new workbook (or load an existing one)
        Workbook wb = new Workbook(); // replace with new Workbook("input.xlsm") to load a file

        // Add a sample checkbox to illustrate detection (remove this block in real scenario)
        int cbIndex = wb.Worksheets[0].CheckBoxes.Add(10, 10, 20, 100);
        wb.Worksheets[0].CheckBoxes[cbIndex].Text = "Sample";

        // Check for form controls before assigning macros
        if (ContainsFormControls(wb))
        {
            Console.WriteLine("Workbook contains form controls. Macros will not be assigned.");
        }
        else
        {
            // Enable macros (or assign macro-related settings here)
            wb.Settings.EnableMacros = true;
            Console.WriteLine("No form controls detected. Macros enabled.");
        }

        // Save the workbook
        wb.Save("ResultWorkbook.xlsx");
    }
}

// Entry point for testing
class Program
{
    static void Main()
    {
        FormControlChecker.RunExample();
    }
}
