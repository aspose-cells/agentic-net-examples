using System;
using Aspose.Cells;

namespace AsposeCellsTsvLoadExample
{
    class Program
    {
        static void Main()
        {
            // Path to the TSV file to be loaded
            string tsvFilePath = "data.tsv";

            // Create load options for a TSV (tab‑separated) file
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Tsv);
            loadOptions.Separator = '\t';               // Tab character as delimiter
            loadOptions.ConvertNumericData = true;      // Convert numeric strings to numbers
            loadOptions.ConvertDateTimeData = true;     // Convert date strings to DateTime values

            // Load the TSV file into a Workbook instance using the specified options
            Workbook workbook = new Workbook(tsvFilePath, loadOptions);

            // Example: read a value from the first worksheet to verify the load
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("First cell value: " + sheet.Cells["A1"].StringValue);
        }
    }
}