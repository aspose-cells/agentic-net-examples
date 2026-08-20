// Title: Aspose.Cells .NET: Define a Named Range for an Entire Column and Compute Its Average
// Description: Demonstrates how to create a new workbook, fill column B with numbers, define a named range that references the whole column using the EntireColumn property, apply an =AVERAGE formula that points to the named range, calculate the workbook, retrieve the result, and save the file.
// Keywords: Aspose.Cells | named range entire column | C# average formula | EntireColumn property | Aspose.Cells calculate formulas | create named range .NET | column average Aspose | Excel named range programmatically | Aspose.Cells workbook calculation
// Common Searches: Aspose.Cells create named range for whole column | How to use EntireColumn property in Aspose.Cells | Calculate average of a column using named range Aspose.Cells | Set formula with named range in Aspose.Cells C# | Reference entire column in named range Aspose.Cells
// Developer Intent: Create a column‑wide named range and use it in an AVERAGE formula with Aspose.Cells for .NET.
// Use Cases: Define a named range that covers column B and calculate its average in cell C1. | Reuse the column‑wide named range across multiple worksheets for aggregate operations such as SUM, COUNT, or AVERAGE. | Programmatically generate reports where column totals are needed without hard‑coding cell references.
// AI Prompts: Show me C# code that creates a named range for an entire column and uses it in an AVERAGE formula with Aspose.Cells. | How can I reference a whole column in a named range and apply SUM, COUNT, or AVERAGE functions using Aspose.Cells for .NET? | Provide an example of using the EntireColumn property to define a column‑wide named range and calculate its average in Aspose.Cells.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a new workbook, fill column B with numbers, define a named range that references the whole column using the EntireColumn property, apply an =AVERAGE formula that points to the named range, calculate the workbook, retrieve the result, and save the file.
    public class NamedRangeEntireColumnAverage
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate some sample numeric data in column B (index 1)
                for (int row = 0; row < 5; row++)
                {
                    cells[row, 1].PutValue(row + 1); // Values 1,2,3,4,5
                }

                // Create a range that starts at B1 (row 0, column 1) with a single cell
                AsposeRange singleCellRange = cells.CreateRange(0, 1, 1, 1);

                // Get the entire column that contains the range (column B)
                AsposeRange entireColumn = singleCellRange.EntireColumn;

                // Define a named range that refers to the whole column B
                int nameIndex = workbook.Worksheets.Names.Add("MyColumn");
                // RefersTo must start with '=' and use absolute column reference
                workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$B:$B";

                // Use the named range in a formula to calculate the average
                cells["C1"].Formula = "=AVERAGE(MyColumn)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Retrieve and display the calculated average
                double average = cells["C1"].DoubleValue;
                Console.WriteLine("Average of MyColumn: " + average);

                // Save the workbook (the file will contain the data and the formula result)
                workbook.Save("NamedRangeEntireColumnAverage.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            NamedRangeEntireColumnAverage.Run();
        }
    }
}
