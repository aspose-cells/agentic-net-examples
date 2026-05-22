using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    public class GetTableDisplayNameDemo
    {
        // Entry point for the application
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Choose a cell that belongs to a table (ensure A1 is inside a table)
                Cell cell = worksheet.Cells["A1"];

                // Retrieve the table (ListObject) that the cell belongs to
                ListObject table = cell.GetTable();

                if (table != null)
                {
                    // Read and output the display name of the table
                    Console.WriteLine("Table Display Name: " + table.DisplayName);
                }
                else
                {
                    Console.WriteLine("The specified cell does not belong to any table.");
                }

                // Optionally save the workbook if any changes were made
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a message
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}