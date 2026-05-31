using System;
using System.IO;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsNumericConversionDemo
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and get the first sheet cells
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // create workbook
            Cells cells = workbook.Worksheets[0].Cells;            // get cells collection

            // -------------------------------------------------
            // 2. Populate cells with textual representations of numbers
            //    - Use PutValue(string, true) to convert while inserting
            //    - Use PutValue(string) to keep as string (no conversion)
            // -------------------------------------------------
            cells["A1"].PutValue("123.45", true);   // converted to double
            cells["A2"].PutValue("678", true);      // converted to int (double internally)
            cells["A3"].PutValue("2023-05-15", true); // converted to DateTime
            cells["A4"].PutValue("NotANumber");     // remains string

            // -------------------------------------------------
            // 3. Demonstrate conversion of remaining string numbers
            //    (if any were inserted without conversion flag)
            // -------------------------------------------------
            cells["B1"].PutValue("987.65"); // inserted as string
            cells["B2"].PutValue("01/01/2022"); // inserted as string
            cells["B3"].PutValue("ABC"); // non‑numeric string

            // Convert all possible string values in the worksheet to numeric/date types
            cells.ConvertStringToNumericValue();

            // -------------------------------------------------
            // 4. Output the resulting cell values and their types
            // -------------------------------------------------
            Console.WriteLine("A1: Value = {0}, Type = {1}", cells["A1"].Value, cells["A1"].Type);
            Console.WriteLine("A2: Value = {0}, Type = {1}", cells["A2"].Value, cells["A2"].Type);
            Console.WriteLine("A3: Value = {0}, Type = {1}", cells["A3"].Value, cells["A3"].Type);
            Console.WriteLine("A4: Value = {0}, Type = {1}", cells["A4"].Value, cells["A4"].Type);
            Console.WriteLine("B1: Value = {0}, Type = {1}", cells["B1"].Value, cells["B1"].Type);
            Console.WriteLine("B2: Value = {0}, Type = {1}", cells["B2"].Value, cells["B2"].Type);
            Console.WriteLine("B3: Value = {0}, Type = {1}", cells["B3"].Value, cells["B3"].Type);

            // -------------------------------------------------
            // 5. Example of loading CSV data with automatic numeric conversion
            // -------------------------------------------------
            string csvData = "ID,Price,Date\n1,19.99,2023-01-01\n2,\"24.50\",\"2023-02-15\"\n3,ABC,NotADate";
            using (MemoryStream ms = new MemoryStream(Encoding.UTF8.GetBytes(csvData)))
            {
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
                {
                    ConvertNumericData = true,      // convert numeric strings
                    ConvertDateTimeData = true      // convert date strings
                };

                Workbook csvWorkbook = new Workbook(ms, loadOptions);
                Cells csvCells = csvWorkbook.Worksheets[0].Cells;

                Console.WriteLine("\nCSV Load - Cell B2 (Price) Type: " + csvCells["B2"].Type);
                Console.WriteLine("CSV Load - Cell C2 (Date) Type: " + csvCells["C2"].Type);
                Console.WriteLine("CSV Load - Cell B3 (Non‑numeric) Type: " + csvCells["B3"].Type);
            }

            // -------------------------------------------------
            // 6. Save the workbook (using the provided save method)
            // -------------------------------------------------
            workbook.Save("NumericConversionResult.xlsx");
        }
    }
}