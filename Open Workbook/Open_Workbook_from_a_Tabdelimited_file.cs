using System;
using Aspose.Cells;

namespace AsposeCellsTabDelimitedDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the tab‑delimited (TSV) file
            string tsvPath = "sample.tsv";

            // Create load options specifying the file is Tab‑Separated Values
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);

            // Open the workbook using the file path and the load options
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Access the first worksheet to demonstrate that the file was loaded
            Worksheet sheet = workbook.Worksheets[0];

            // Output the value of the first cell (A1) to the console
            Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

            // (Optional) Save the workbook as an XLSX file to verify conversion
            workbook.Save("ConvertedFromTsv.xlsx", SaveFormat.Xlsx);
        }
    }
}