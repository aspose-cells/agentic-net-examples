// Title: Read the third sparkline value from the second sparkline group in an Excel file using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# read value of a sparkline in a specific sparkline group | how to get the data range of the third sparkline in the second group with Aspose.Cells | C# code to retrieve source cell value of a sparkline using Aspose.Cells | extract sparkline source cell value from an Excel workbook with Aspose.Cells .NET | log sparkline value from second sparkline group using Aspose.Cells API
// Tags: read sparkline data range Aspose.Cells | access third sparkline in sparkline group C# | retrieve source cell value from sparkline Aspose.Cells | log sparkline value .NET | excel sparkline group manipulation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, verifies that at least two sparkline groups exist, selects the third sparkline from the second group, reads its source cell via the DataRange property, and writes the cell's value to the console using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Ensure the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that at least two sparkline groups exist
            if (sheet.SparklineGroups.Count < 2)
            {
                Console.WriteLine("The worksheet contains fewer than two sparkline groups.");
                return;
            }

            // Get the second sparkline group (index 1 because collections are zero‑based)
            var secondGroup = sheet.SparklineGroups[1];

            // Verify that the second group has at least three sparklines
            if (secondGroup.SparklineCollection.Count < 3)
            {
                Console.WriteLine("The second sparkline group contains fewer than three sparklines.");
                return;
            }

            // Retrieve the third sparkline in the second group (index 2)
            var thirdSparkline = secondGroup.SparklineCollection[2];

            // The DataRange property holds the address of the source cell (e.g., "C5")
            string dataRange = thirdSparkline.DataRange;

            // Access the source cell directly by its address
            Cell sourceCell = sheet.Cells[dataRange];

            // Read the cell value
            object value = sourceCell.Value;

            // Log the value
            Console.WriteLine($"Value of the third sparkline in the second group: {value}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
