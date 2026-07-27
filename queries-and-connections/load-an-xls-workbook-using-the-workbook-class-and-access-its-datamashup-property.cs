using System;
using Aspose.Cells;
using Aspose.Cells.QueryTables;

class Program
{
    static void Main()
    {
        // Path to the existing XLS workbook
        string sourcePath = "input.xls";

        // Load the workbook from the file using the Workbook(string) constructor
        Workbook workbook = new Workbook(sourcePath);

        // Access the DataMashup property
        DataMashup mashup = workbook.DataMashup;

        // Demonstrate usage of the DataMashup object
        if (mashup != null)
        {
            var powerQueryFormulas = mashup.PowerQueryFormulas;
            Console.WriteLine($"Number of Power Query formulas: {powerQueryFormulas.Count}");
        }
        else
        {
            Console.WriteLine("DataMashup is null.");
        }

        // Save the workbook to a new file (optional)
        string destPath = "output.xlsx";
        workbook.Save(destPath);
    }
}