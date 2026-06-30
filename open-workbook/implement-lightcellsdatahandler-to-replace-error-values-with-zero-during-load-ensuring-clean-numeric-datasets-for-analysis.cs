using System;
using System.IO;
using Aspose.Cells;

namespace LightCellsErrorHandlerDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Prepare a sample CSV with an error formula (e.g., division by zero)
            File.WriteAllText("input.csv", "A,B,C\n10,0,=A2/B2\n5,2,=A3/B3");

            // Create load options and assign the custom LightCellsDataHandler
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            loadOptions.LightCellsDataHandler = new ReplaceErrorWithZeroHandler();

            // Load workbook using the custom handler
            Workbook workbook = new Workbook("input.csv", loadOptions);

            // Save the processed workbook
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook loaded and errors replaced with zero. Saved as output.xlsx");
        }
    }

    // Custom LightCellsDataHandler that replaces error cells with zero
    public class ReplaceErrorWithZeroHandler : LightCellsDataHandler
    {
        public bool StartSheet(Worksheet sheet)
        {
            // Process all sheets
            return true;
        }

        public bool StartRow(int rowIndex)
        {
            // Process all rows
            return true;
        }

        public bool StartCell(int columnIndex)
        {
            // Process all cells
            return true;
        }

        public bool ProcessCell(Cell cell)
        {
            // If the cell contains an error, replace it with numeric zero
            if (cell.Type == CellValueType.IsError)
            {
                cell.PutValue(0);
            }
            return true; // Continue processing
        }

        public bool ProcessRow(Row row)
        {
            // No additional row processing needed
            return true;
        }
    }
}