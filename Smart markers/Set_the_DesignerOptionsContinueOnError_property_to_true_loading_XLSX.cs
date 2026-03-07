using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create LoadOptions for XLSX format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Load the workbook with the specified options
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}