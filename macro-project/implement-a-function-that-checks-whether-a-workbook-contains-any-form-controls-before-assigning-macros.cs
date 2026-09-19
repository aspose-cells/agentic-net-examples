// Title: Check if an Excel workbook contains form controls before assigning macros with Aspose.Cells (C#)
// AI Prompts: Write a C# method that loads an Excel file using Aspose.Cells, iterates all worksheets, and returns true when any shape of type Button, CheckBox, ComboBox, ListBox, OptionButton, or RadioButton is encountered. | Enhance the detection method to log the worksheet name and the exact form‑control type for each matching shape found. | Add robust try‑catch handling so the method safely returns false and records the exception message if shape enumeration fails.
// Common Searches: how to programmatically detect Excel form controls using Aspose.Cells in C# | Aspose.Cells C# enumerate worksheet shapes to find buttons or checkboxes | verify workbook has no form controls before attaching macros with Aspose.Cells | C# sample code to check for radio buttons or combo boxes in an Excel file using Aspose.Cells | detect and list form control types in an .xlsx workbook with Aspose.Cells API
// Tags: Aspose.Cells detect form controls | C# iterate worksheet shape collection | Excel macro pre‑validation Aspose.Cells | shape type filtering Aspose.Cells | check workbook for button checkbox shapes

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    // A helper method iterates through each worksheet's ShapeCollection in an Aspose.Cells Workbook, checks each shape's Type against common form‑control types (Button, CheckBox, ComboBox, ListBox, OptionButton, RadioButton), and returns true if any are found, with null‑workbook checks and exception handling.
    public static class WorkbookHelper
    {
        // Checks if the workbook contains any form controls (e.g., button, checkbox, combobox, etc.)
        public static bool ContainsFormControls(Workbook workbook)
        {
            if (workbook == null)
                return false;

            try
            {
                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the collection of shapes on the worksheet
                    ShapeCollection shapes = sheet.Shapes;

                    // Iterate through each shape
                    foreach (Shape shape in shapes)
                    {
                        // Identify form controls by their type name (avoids direct enum reference)
                        string typeName = shape.Type.ToString();

                        if (typeName == "Button" ||
                            typeName == "CheckBox" ||
                            typeName == "ComboBox" ||
                            typeName == "ListBox" ||
                            typeName == "OptionButton" ||
                            typeName == "RadioButton")
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while checking form controls: {ex.Message}");
            }

            // No form controls found
            return false;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(filePath);

                // Check for form controls
                bool hasControls = WorkbookHelper.ContainsFormControls(workbook);
                Console.WriteLine(hasControls
                    ? "The workbook contains form controls."
                    : "The workbook does not contain any form controls.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
