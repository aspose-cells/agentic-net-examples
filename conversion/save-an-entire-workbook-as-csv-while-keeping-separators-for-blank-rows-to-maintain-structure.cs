using System;
using System.Text;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get its first worksheet's cells
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Populate some data with intentional blank rows in between
        cells[0, 0].PutValue("a");
        cells[0, 1].PutValue("b");
        // rows 1 and 2 remain blank
        cells[3, 0].PutValue("c");
        cells[4, 1].PutValue("d");

        // Configure CSV save options to keep separators for blank rows
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Encoding = Encoding.UTF8,
            Separator = ',',               // standard comma delimiter
            KeepSeparatorsForBlankRow = true // output separators even for empty rows
        };

        // Save the entire workbook as a CSV file using the configured options
        workbook.Save("output.csv", csvOptions);
    }
}