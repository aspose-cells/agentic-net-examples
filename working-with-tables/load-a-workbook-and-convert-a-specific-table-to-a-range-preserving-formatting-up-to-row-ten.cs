// Title: Convert a named Excel table to a regular range while preserving formatting with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# convert ListObject Table1 to range preserving formatting | How to change an Excel table to a normal range using Aspose.Cells .NET | C# code to convert a named Excel table to a range without losing cell styles | Aspose.Cells example for converting a worksheet table to a range and saving the workbook
// Tags: Aspose.Cells ListObject conversion | preserve formatting during Excel table conversion | load workbook modify ListObject .NET | Excel table to range conversion C# | Aspose.Cells table to range example

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The sample loads 'input.xlsx', checks its existence, accesses the first worksheet, retrieves the ListObject named 'Table1', converts the table into a regular range while keeping all data and formatting, and saves the result as 'output.xlsx', with error handling for missing files or tables.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or specify by name if needed)
            Worksheet worksheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

            // Retrieve the table (ListObject) by its name.
            // Ensure the table name matches the one defined in the workbook.
            ListObject table = worksheet.ListObjects["Table1"];
            if (table == null)
            {
                Console.WriteLine("Error: Table \"Table1\" not found in the worksheet.");
                return;
            }

            // Convert the table to a normal range while preserving data and formatting
            table.ConvertToRange();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
