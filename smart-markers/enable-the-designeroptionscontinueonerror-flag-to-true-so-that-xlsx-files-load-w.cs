using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Set load options (default behavior will be used).
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Load the workbook with the specified options.
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook to a new file.
        workbook.Save("output.xlsx");
    }
}