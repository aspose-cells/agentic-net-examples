using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsFormControlCheck
{
    public class MacroHelper
    {
        // Checks if the workbook contains any form controls (CheckBoxes or ActiveX controls)
        public static bool ContainsFormControls(Workbook workbook)
        {
            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Check for legacy form controls (CheckBox collection)
                if (sheet.CheckBoxes.Count > 0)
                {
                    return true;
                }

                // Check for ActiveX controls embedded as shapes
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.ActiveXControl != null)
                    {
                        return true;
                    }
                }
            }

            // No form controls found
            return false;
        }

        // Example method that assigns macros only if no form controls are present
        public static void AssignMacroIfNoFormControls(string inputPath, string outputPath)
        {
            // Load the workbook (create/load rule)
            Workbook wb = new Workbook(inputPath);

            // Verify presence of form controls before enabling macros
            if (ContainsFormControls(wb))
            {
                Console.WriteLine("Workbook contains form controls. Macros will not be assigned.");
                // Optionally, you could remove macros or take other actions here
            }
            else
            {
                Console.WriteLine("No form controls detected. Enabling macros.");
                // Enable macros for the workbook
                wb.Settings.EnableMacros = true;

                // Example: add a simple VBA module (requires macro-enabled format)
                // Ensure the workbook is saved as a macro-enabled file
                wb.Save(outputPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook saved with macros enabled at: {outputPath}");
                return;
            }

            // Save the workbook without macros (or with macros removed)
            wb.RemoveMacro(); // Ensure macros are removed if they existed
            wb.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved without macros at: {outputPath}");
        }

        // Entry point for demonstration
        public static void Main()
        {
            string sourceFile = "InputWorkbook.xlsx";   // Path to the source workbook
            string resultFile = "ResultWorkbook.xlsm";   // Desired output path

            AssignMacroIfNoFormControls(sourceFile, resultFile);
        }
    }
}