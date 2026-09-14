// Title: Apply a custom cell style to smart marker generated cells in an Aspose.Cells workbook (C#)
// AI Prompts: Write C# code that loads a workbook containing smart markers, processes them with WorkbookDesigner, creates a Style with a solid light‑gray background and centered text, and applies it to the populated cells. | Show how to configure a StyleFlag to apply all formatting attributes to a worksheet or range after smart marker processing using Aspose.Cells. | Demonstrate saving the workbook to an Xlsx file after applying the custom style to smart marker output cells.
// Common Searches: Aspose.Cells C# apply formatting after WorkbookDesigner.Process | How to style smart marker output cells in a generated Excel file using Aspose.Cells | C# set background color and alignment for smart marker populated range Aspose.Cells | Apply a custom style to entire worksheet after smart markers are processed in Aspose.Cells | Use StyleFlag to apply all style properties to smart marker cells in C#
// Tags: apply style after smart marker processing Aspose.Cells | WorkbookDesigner custom cell formatting C# | StyleFlag all properties Aspose.Cells | set background color smart marker cells | uniform worksheet styling Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerStyling
{
    // The example loads a template workbook with smart markers, binds a list of Employee objects as the data source, processes the markers using WorkbookDesigner, creates a light‑gray centered style, applies it to the entire worksheet with a StyleFlag set to All, and saves the styled workbook as StyledOutput.xlsx.
    class Program
    {
        static void Main()
        {
            // Load a template workbook that contains smart markers
            Workbook workbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare sample data source
            var employees = new List<Employee>
            {
                new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                new Employee { Name = "Jane Smith", Age = 28, Department = "Marketing" }
            };

            // Set up the WorkbookDesigner with the loaded workbook
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Assign the data source to the smart marker name used in the template
            designer.SetDataSource("Employees", employees);

            // Process all smart markers in the workbook
            designer.Process();

            // -----------------------------------------------------------------
            // After processing, apply a custom style to the cells that were
            // populated by smart markers (or to the whole worksheet to keep
            // formatting consistent).
            // -----------------------------------------------------------------

            // Create a style: light gray background with centered text
            Style customStyle = workbook.CreateStyle();
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.ForegroundColor = System.Drawing.Color.LightGray;
            customStyle.HorizontalAlignment = TextAlignmentType.Center;
            customStyle.VerticalAlignment = TextAlignmentType.Center;
            customStyle.Font.Name = "Calibri";
            customStyle.Font.Size = 11;

            // Define which style attributes should be applied
            StyleFlag flag = new StyleFlag
            {
                All = true               // Apply all formatting properties defined above
            };

            // Apply the style to the entire worksheet (or you could limit it to a range)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells.ApplyStyle(customStyle, flag);

            // Save the resulting workbook
            workbook.Save("StyledOutput.xlsx", SaveFormat.Xlsx);
        }

        // Simple POCO class used as a data source for the smart markers
        public class Employee
        {
            public string Name { get; set; }
            public int Age { get; set; }
            public string Department { get; set; }
        }
    }
}
