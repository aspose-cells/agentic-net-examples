// Title: Read raw strings from a worksheet range and concatenate them into a single cell using Aspose.Cells for .NET
// Common Searches: how to read raw string values from a range of cells with Aspose.Cells C# | concatenate values from A1:A10 and place result in B1 using Aspose.Cells | Aspose.Cells example for merging cell text into one cell in .NET | C# code to combine multiple Excel cells into a summary cell with Aspose.Cells | using StringBuilder to concatenate Excel range values in Aspose.Cells
// Tags: read raw string values Aspose.Cells range | concatenate cell values Aspose.Cells C# | write summary cell Aspose.Cells | StringBuilder Excel range Aspose.Cells | save workbook after concatenation Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example loads (or creates) a workbook, reads the raw string values from a specified range (e.g., A1:A10) on the first worksheet, concatenates them with a StringBuilder, writes the combined text to a summary cell (e.g., B1), and saves the workbook, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load an existing workbook if the file exists; otherwise create a new workbook
            Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to read (e.g., A1:A10)
            string sourceRange = "A1:A10";

            // Get the cells collection
            Cells cells = sheet.Cells;

            // Retrieve the Aspose.Cells range object
            Aspose.Cells.Range range = cells.CreateRange(sourceRange);

            // Use StringBuilder for efficient concatenation
            StringBuilder concatenated = new StringBuilder();

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Get the raw string value of the cell (empty string if null)
                string cellText = cell.StringValue ?? string.Empty;

                // Append to the result
                concatenated.Append(cellText);
            }

            // Write the concatenated result to a summary cell (e.g., B1)
            Cell summaryCell = cells["B1"];
            summaryCell.PutValue(concatenated.ToString());

            // Save the workbook to the desired output path
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log or handle exceptions as needed
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
