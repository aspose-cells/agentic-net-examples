// Title: Get an Excel Table's DisplayName with Cell.GetTable in Aspose.Cells for .NET (C#)
// Description: Loads a workbook, selects a cell inside a table, uses Cell.GetTable to obtain the ListObject, reads its DisplayName property, handles a possible null result, prints the name, and saves the file unchanged.
// Keywords: Aspose.Cells | Cell.GetTable | ListObject DisplayName | C# Excel table name | retrieve table display name | read Excel table name .NET | Aspose.Cells example C# | Excel table metadata | Aspose.Cells GetTable demo
// Common Searches: Aspose.Cells get table display name C# | Cell.GetTable return ListObject name | how to read Excel table name with Aspose.Cells | retrieve ListObject DisplayName from a cell | C# example for getting Excel table name
// Developer Intent: Extract the display name of the table that contains a specific cell in an Excel workbook.
// Use Cases: Log table names before processing rows for auditing. | Validate that a cell belongs to the expected table by comparing its DisplayName. | Generate documentation that lists all table names in a workbook.
// AI Prompts: Show how to safely check for a null ListObject when using Cell.GetTable in Aspose.Cells. | Provide C# code to iterate over all tables in a worksheet and output each DisplayName. | Explain how to rename a table's DisplayName after retrieving it with Cell.GetTable.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Loads a workbook, selects a cell inside a table, uses Cell.GetTable to obtain the ListObject, reads its DisplayName property, handles a possible null result, prints the name, and saves the file unchanged.
    public class GetTableDisplayNameDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Choose a cell that belongs to a table (e.g., A2)
                Cell cellInTable = worksheet.Cells["A2"];

                // Retrieve the table (ListObject) that the cell belongs to
                ListObject table = cellInTable.GetTable();

                if (table != null)
                {
                    // Read and display the table's display name
                    string displayName = table.DisplayName;
                    Console.WriteLine("Table Display Name: " + displayName);
                }
                else
                {
                    Console.WriteLine("The specified cell is not part of any table.");
                }

                // Save the workbook (unchanged) to a new file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
