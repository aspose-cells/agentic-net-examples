using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get its first worksheet's cells
            Workbook wb = new Workbook();
            Cells cells = wb.Worksheets[0].Cells;

            // Add data with intentional blank rows (rows 1 and 2 are left empty)
            cells[0, 0].PutValue("a");
            cells[0, 1].PutValue("b");
            cells[3, 0].PutValue("c");
            cells[4, 1].PutValue("d");

            // Configure CSV save options:
            // - Use comma as separator
            // - Set encoding to UTF-8
            // - Keep separators for blank rows so empty rows are represented as ","
            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                Separator = ',',
                Encoding = Encoding.UTF8,
                KeepSeparatorsForBlankRow = true
            };

            // Save the workbook as a CSV file with the specified options
            wb.Save("output.csv", saveOptions);
        }
    }
}