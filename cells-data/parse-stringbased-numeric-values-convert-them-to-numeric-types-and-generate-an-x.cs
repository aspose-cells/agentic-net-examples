using System;
using Aspose.Cells;

namespace AsposeCellsStringToNumericDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Sample data: string representations of numbers, dates and non‑numeric text
            string[,] data = new string[,]
            {
                { "123.45", "2021-06-20", "ABC" },
                { "678",    "15/08/2020", "DEF" },
                { "9.81",   "invalid",   "GHI" }
            };

            // Populate the worksheet with the string data
            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    cells[row, col].PutValue(data[row, col]); // store as string initially
                }
            }

            // Convert all possible string values to their numeric (or date) equivalents
            cells.ConvertStringToNumericValue();

            // Save the workbook as an XLSX file
            workbook.Save("StringToNumericOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}