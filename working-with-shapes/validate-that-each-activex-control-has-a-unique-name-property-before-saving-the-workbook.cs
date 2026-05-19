using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

namespace AsposeCellsActiveXValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a few ActiveX controls with intentional duplicate names
            Shape shape1 = sheet.Shapes.AddActiveXControl(ControlType.CheckBox, 2, 0, 2, 0, 100, 30);
            shape1.Name = "MyControl";
            CheckBoxActiveXControl cb1 = (CheckBoxActiveXControl)shape1.ActiveXControl;
            cb1.Caption = "First";

            Shape shape2 = sheet.Shapes.AddActiveXControl(ControlType.CommandButton, 5, 0, 5, 0, 100, 30);
            shape2.Name = "MyControl"; // Duplicate name
            ActiveXControl btn1 = shape2.ActiveXControl;
            btn1.IsEnabled = true;

            Shape shape3 = sheet.Shapes.AddActiveXControl(ControlType.ComboBox, 8, 0, 8, 0, 100, 30);
            shape3.Name = "UniqueControl";
            ComboBoxActiveXControl cbx1 = (ComboBoxActiveXControl)shape3.ActiveXControl;
            cbx1.ListFillRange = "A1:A5";

            // Validate that each ActiveX control has a unique name
            EnsureUniqueActiveXNames(sheet);

            // Save the workbook (validation already performed)
            workbook.Save("ValidatedActiveXControls.xlsx");
        }

        /// <summary>
        /// Ensures all ActiveX controls on the given worksheet have unique names.
        /// If a duplicate is found, the name is modified by appending a numeric suffix.
        /// </summary>
        /// <param name="worksheet">Worksheet to inspect.</param>
        static void EnsureUniqueActiveXNames(Worksheet worksheet)
        {
            // Keep track of names we have already seen
            HashSet<string> existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // Counter for generating new unique names
            int duplicateCounter = 1;

            // Iterate through all shapes on the worksheet
            foreach (Shape shape in worksheet.Shapes)
            {
                // Only process shapes that host an ActiveX control
                if (shape.ActiveXControl != null)
                {
                    string originalName = shape.Name;

                    // If the name is null or empty, assign a default unique name
                    if (string.IsNullOrWhiteSpace(originalName))
                    {
                        string generatedName;
                        do
                        {
                            generatedName = $"ActiveX_{duplicateCounter++}";
                        } while (!existingNames.Add(generatedName));

                        shape.Name = generatedName;
                        continue;
                    }

                    // If the name already exists, modify it until it becomes unique
                    if (!existingNames.Add(originalName))
                    {
                        string newName;
                        do
                        {
                            newName = $"{originalName}_{duplicateCounter++}";
                        } while (!existingNames.Add(newName));

                        shape.Name = newName;
                    }
                }
            }
        }
    }
}