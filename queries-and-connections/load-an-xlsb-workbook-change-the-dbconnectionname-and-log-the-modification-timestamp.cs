using System;
using Aspose.Cells;
using Aspose.Cells.ExternalConnections;

class Program
{
    static void Main()
    {
        // Path to the source XLSB file
        string inputPath = "input.xlsb";

        // Path for the modified workbook
        string outputPath = "output.xlsb";

        // Load the XLSB workbook
        Workbook workbook = new Workbook(inputPath);

        // Change the Name of each DBConnection in the workbook
        foreach (ExternalConnection conn in workbook.DataConnections)
        {
            if (conn is DBConnection dbConn)
            {
                // Set a new name (example uses a timestamp to ensure uniqueness)
                dbConn.Name = "ModifiedConnection_" + DateTime.Now.Ticks;
            }
        }

        // Log the modification timestamp
        Console.WriteLine($"DBConnection name(s) modified at: {DateTime.Now:O}");

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}