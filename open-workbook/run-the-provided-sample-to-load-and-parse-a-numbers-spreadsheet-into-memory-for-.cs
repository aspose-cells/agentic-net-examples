using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Numbers;

namespace NumbersLoadExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Apple Numbers file (adjust as needed)
            string numbersFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.numbers");

            Workbook workbook;

            if (File.Exists(numbersFilePath))
            {
                // Create NumbersLoadOptions and configure as needed
                NumbersLoadOptions loadOptions = new NumbersLoadOptions
                {
                    // Load each table as a separate worksheet
                    LoadTableType = LoadNumbersTableType.OneTablePerSheet,
                    // Keep unparsed data for potential later use
                    KeepUnparsedData = true
                };

                // Load the Numbers file into a Workbook using the specified options
                workbook = new Workbook(numbersFilePath, loadOptions);
            }
            else
            {
                // If the .numbers file is not found, create a new workbook with sample data
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells sheetCells = sheet.Cells;
                sheetCells["A1"].PutValue("Sample");
                sheetCells["B1"].PutValue(123);
                sheetCells["A2"].PutValue(DateTime.Now);
            }

            // Access the first worksheet for demonstration
            Worksheet worksheet = workbook.Worksheets[0];
            Cells worksheetCells = worksheet.Cells;

            // Print the value of cell A1 (if any) to verify successful load
            Console.WriteLine("Cell A1 value: " + worksheetCells["A1"].StringValue);

            // Iterate through used rows and columns
            int maxRow = worksheetCells.MaxDataRow;
            int maxColumn = worksheetCells.MaxDataColumn;

            Console.WriteLine("Data in the first worksheet:");
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Console.Write(worksheetCells[row, col].StringValue + "\t");
                }
                Console.WriteLine();
            }

            // (Optional) Save the workbook to another format for further processing
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}