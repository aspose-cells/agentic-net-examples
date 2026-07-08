using System;
using Aspose.Cells;

namespace AsposeCellsCellTypeCounter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate the worksheet with sample data of various types
            cells["A1"].PutValue(123);                     // Numeric
            cells["A2"].PutValue("Hello World");           // String
            cells["A3"].PutValue(DateTime.Now);            // DateTime
            cells["A4"].PutValue(true);                    // Boolean
            cells["A5"].PutValue(null);                    // Null
            cells["A6"].PutValue("=1/0");                  // Error (division by zero)

            // Counters for each cell value type
            int numericCount = 0;
            int stringCount = 0;
            int dateTimeCount = 0;
            int boolCount = 0;
            int nullCount = 0;
            int errorCount = 0;
            int unknownCount = 0;

            // Iterate through all instantiated cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Skip cells that have never been instantiated (value is null and type is Unknown)
                // The Cells collection enumerates only instantiated cells, so we can directly evaluate the type.
                switch (cell.Type)
                {
                    case CellValueType.IsNumeric:
                        numericCount++;
                        break;
                    case CellValueType.IsString:
                        stringCount++;
                        break;
                    case CellValueType.IsDateTime:
                        dateTimeCount++;
                        break;
                    case CellValueType.IsBool:
                        boolCount++;
                        break;
                    case CellValueType.IsNull:
                        nullCount++;
                        break;
                    case CellValueType.IsError:
                        errorCount++;
                        break;
                    default:
                        unknownCount++;
                        break;
                }
            }

            // Output the results
            Console.WriteLine("Cell type counts:");
            Console.WriteLine($"Numeric   : {numericCount}");
            Console.WriteLine($"String    : {stringCount}");
            Console.WriteLine($"DateTime  : {dateTimeCount}");
            Console.WriteLine($"Boolean   : {boolCount}");
            Console.WriteLine($"Null      : {nullCount}");
            Console.WriteLine($"Error     : {errorCount}");
            Console.WriteLine($"Unknown   : {unknownCount}");

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("CellTypeCountDemo.xlsx");
        }
    }
}