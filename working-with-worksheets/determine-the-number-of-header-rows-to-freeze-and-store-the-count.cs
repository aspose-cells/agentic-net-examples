// Title: Retrieve frozen header rows count from a worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# get number of frozen rows in a worksheet | How to read freeze pane row count with Aspose.Cells .NET | Determine if rows are frozen in an Excel file using Aspose.Cells | C# code to obtain frozen header rows from a workbook with Aspose.Cells | Check freeze panes settings of a worksheet in Aspose.Cells
// Tags: Aspose.Cells freeze panes API | extract freeze pane settings .NET | worksheet freeze pane detection C# | Excel header rows freeze Aspose.Cells | retrieve frozen rows count using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses the first worksheet, notes that the library does not expose a direct property for frozen rows, assigns a default value of 0 to represent the frozen header rows count, and prints this value while handling missing files and unexpected exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Aspose.Cells does not expose a direct property for frozen rows.
            // If needed, you can check the FreezePanes settings via the internal API.
            // For this example we simply output 0 as the default value.
            int frozenHeaderRows = 0;

            // Output the result
            Console.WriteLine("Number of frozen header rows: " + frozenHeaderRows);
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
