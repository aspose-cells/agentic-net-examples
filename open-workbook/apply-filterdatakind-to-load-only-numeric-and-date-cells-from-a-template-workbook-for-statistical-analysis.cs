using System;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    class Program
    {
        static void Main()
        {
            // Path to the template workbook
            string templatePath = "TemplateWorkbook.xlsx";

            // Create LoadOptions and set a LoadFilter that loads only numeric (including date/time) cells
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LoadFilter = new LoadFilter(LoadDataFilterOptions.CellNumeric);

            // Load the workbook using the specified load options
            Workbook workbook = new Workbook(templatePath, loadOptions);

            // At this point, only cells with numeric or date values are loaded.
            // Perform any statistical analysis here, e.g., iterate through worksheets and cells.

            // Example: Count numeric cells in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            int numericCellCount = 0;
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Type == CellValueType.IsNumeric) // includes dates
                {
                    numericCellCount++;
                }
            }
            Console.WriteLine($"Numeric/Date cells loaded: {numericCellCount}");

            // Save the filtered workbook (optional, for verification)
            workbook.Save("FilteredWorkbook.xlsx");
        }
    }
}