// Title: How to concatenate raw string values from an Aspose.Cells range and write the result to a summary cell in C#
// AI Prompts: Write C# code that iterates over an Aspose.Cells Range, extracts each cell's StringValue, concatenates the values with a StringBuilder, and stores the combined text in a target cell. | Show an example of using Aspose.Cells for .NET to read raw strings from a multi‑cell range, merge them into a single string, write the merged string to another cell, and then save the workbook.
// Common Searches: Aspose.Cells C# concatenate values from A1:C2 into D1 | retrieve StringValue from each cell in a range using Aspose.Cells .NET | combine multiple cell strings into one cell with Aspose.Cells API | how to iterate over an Aspose.Cells range and build a single string | save workbook after merging cell text in Aspose.Cells C#
// Tags: concatenate range StringValue Aspose.Cells C# | write merged text to cell Aspose.Cells | Aspose.Cells range iteration example | StringBuilder usage with Aspose.Cells | save workbook after cell concatenation Aspose.Cells

using System;
using System.Text;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsConcatenateExample
{
    // The example creates a workbook, fills cells A1:C2 with text, iterates over the defined range to collect each cell's raw StringValue, concatenates them using StringBuilder, writes the combined string to cell D1, and saves the file as ConcatenatedResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate a sample range with string values
                cells["A1"].PutValue("Hello");
                cells["B1"].PutValue("World");
                cells["C1"].PutValue("!");
                cells["A2"].PutValue("Foo");
                cells["B2"].PutValue("Bar");
                cells["C2"].PutValue("Baz");

                // Define the range from which raw strings will be retrieved
                AsposeRange sourceRange = cells.CreateRange("A1:C2");

                // Concatenate the raw string values of all cells in the range
                StringBuilder concatenated = new StringBuilder();
                foreach (Cell cell in sourceRange)
                {
                    // StringValue returns the formatted text of the cell
                    concatenated.Append(cell.StringValue);
                }

                // Write the concatenated result to a summary cell (e.g., D1)
                cells["D1"].PutValue(concatenated.ToString());

                // Save the workbook to a file
                workbook.Save("ConcatenatedResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
