using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the configuration file that contains the number of rows to freeze
        const string configFile = "freezeConfig.txt";

        // Default to no frozen rows if the config cannot be read
        int frozenRows = 0;

        // Read and parse the configuration value
        if (File.Exists(configFile))
        {
            string text = File.ReadAllText(configFile).Trim();
            int.TryParse(text, out frozenRows);
        }
        else
        {
            Console.WriteLine($"Configuration file '{configFile}' not found. No rows will be frozen.");
        }

        // Create a new workbook (using the standard Aspose.Cells creation pattern)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Apply freeze panes based on the configuration value
        // Freeze at column 0 (first column) and at the row index equal to frozenRows
        // The last two parameters specify how many rows and columns are actually frozen
        if (frozenRows > 0)
        {
            worksheet.FreezePanes(frozenRows, 0, frozenRows, 0);
        }

        // Save the workbook to disk (using the standard Aspose.Cells save pattern)
        const string outputFile = "FrozenRowsDemo.xlsx";
        workbook.Save(outputFile);

        Console.WriteLine($"Workbook saved to '{outputFile}' with {frozenRows} frozen rows.");
    }
}