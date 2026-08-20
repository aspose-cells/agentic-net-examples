// Title: Detect Form Controls in an Excel Workbook and Conditionally Enable Macros with Aspose.Cells for .NET
// Description: Iterates through each worksheet of a loaded Workbook, examines the CheckBoxes collection and Shape objects for ActiveX controls, and returns true if any form control is present. The helper enables macros only when no controls are found, then saves the file.
// Keywords: Aspose.Cells | detect form controls | enable macros conditionally | ActiveX control detection | Excel checkbox check | .NET workbook automation
// Common Searches: Aspose.Cells check for form controls in Excel | Enable macros only if workbook has no ActiveX or checkboxes | How to detect checkboxes in Aspose.Cells .NET | Conditional macro activation based on worksheet shapes
// Developer Intent: Identify whether an Excel workbook contains any form controls and activate macros only when the workbook is free of those controls.
// Use Cases: Validate template files before distribution to prevent macro security warnings caused by embedded controls. | Batch‑process a collection of workbooks, turning on macros only for files without checkboxes or ActiveX objects. | Integrate into CI/CD pipelines to enforce a policy that production Excel files do not mix macros with legacy form controls.
// AI Prompts: Write unit tests for ContainsFormControls covering worksheets with checkboxes, ActiveX shapes, and no controls. | Modify ProcessWorkbook to log the names of sheets that contain form controls before skipping macro enablement. | Extend the helper to detect additional FormControl types such as radio buttons and return a list of detected control types.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Iterates through each worksheet of a loaded Workbook, examines the CheckBoxes collection and Shape objects for ActiveX controls, and returns true if any form control is present. The helper enables macros only when no controls are found, then saves the file.
public class WorkbookMacroHelper
{
    // Checks if any worksheet in the workbook contains form controls
    // (CheckBoxes collection or ActiveX controls).
    public static bool ContainsFormControls(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check for legacy form controls (CheckBoxes)
            if (sheet.CheckBoxes.Count > 0)
                return true;

            // Check for ActiveX controls added as shapes
            foreach (Shape shape in sheet.Shapes)
            {
                if (shape.ActiveXControl != null)
                    return true;
            }
        }
        return false;
    }

    // Loads a workbook, checks for controls, and enables macros only when no form controls are present.
    public static void ProcessWorkbook(string inputPath, string outputPath)
    {
        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from file
            Workbook wb = new Workbook(inputPath);

            // Verify if the workbook contains any form controls
            bool hasControls = ContainsFormControls(wb);
            Console.WriteLine($"Workbook contains form controls: {hasControls}");

            // Assign macros only if there are no form controls
            if (!hasControls)
            {
                wb.Settings.EnableMacros = true;
                Console.WriteLine("Macros have been enabled.");
            }
            else
            {
                Console.WriteLine("Macros were not enabled due to existing form controls.");
            }

            // Save the workbook to the specified output path
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <inputPath> <outputPath>");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];

        WorkbookMacroHelper.ProcessWorkbook(inputPath, outputPath);
    }
}
