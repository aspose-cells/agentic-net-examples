using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.ActiveXControls;

class ListFormControlsWithMacros
{
    // Custom LoadFilter that loads only VBA data from each worksheet
    private class VbaLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only VBA projects for the sheet
            LoadDataFilterOptions = LoadDataFilterOptions.VBA;
        }
    }

    static void Main()
    {
        // Prepare load options with the custom filter
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LoadFilter = new VbaLoadFilter();

        // Load the workbook (only VBA data is loaded)
        Workbook workbook = new Workbook("input.xlsm", loadOptions);

        // Verify that the workbook contains macros
        Console.WriteLine("Workbook HasMacro: " + workbook.HasMacro);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"Worksheet: {sheet.Name}");

            // Iterate through all shapes in the worksheet
            foreach (Shape shape in sheet.Shapes)
            {
                // Check if the shape hosts an ActiveX control (form control)
                if (shape.ActiveXControl != null)
                {
                    // Try to obtain the macro name assigned to the control.
                    // The property may be named "MacroName" or simply "Macro".
                    string macroName = null;
                    var control = shape.ActiveXControl;
                    var type = control.GetType();

                    var macroProp = type.GetProperty("MacroName") ?? type.GetProperty("Macro");
                    if (macroProp != null)
                    {
                        macroName = macroProp.GetValue(control) as string;
                    }

                    // If a macro is assigned, output the control information
                    if (!string.IsNullOrEmpty(macroName))
                    {
                        Console.WriteLine($"  Control: {shape.Name}, Type: {type.Name}, Macro: {macroName}");
                    }
                }
            }
        }
    }
}