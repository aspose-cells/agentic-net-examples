using System;
using Aspose.Cells;

class WorkbookToCsv
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure CSV save options to preserve leading blank rows/columns
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.TrimLeadingBlankRowAndColumn = false; // keep original column structure unchanged

        // Save the workbook as CSV without adding any row numbers
        workbook.Save("output.csv", csvOptions);
    }
}