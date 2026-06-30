using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create LoadOptions and configure it to load only cell values (no formatting).
        LoadOptions loadOptions = new LoadOptions();

        // NOTE: The LoadDataOnly property is not documented in the provided API reference.
        // If your Aspose.Cells version supports it, uncomment the line below.
        // loadOptions.LoadDataOnly = true; // Load only data without formatting.

        // Load the workbook using the configured options.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook (formatting will be default since only data was loaded).
        workbook.Save("output.xlsx");
    }
}