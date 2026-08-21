// Title: Read‑Only Cells Enumerator to Extract All Numeric Values in Aspose.Cells for .NET
// Description: Shows how to create a workbook, fill it with mixed data, optionally convert numeric strings, obtain a read‑only Cells enumerator, iterate without modifying the collection, use Cell.IsNumericValue to identify numbers, collect Cell.DoubleValue into a List<double>, and display the results.
// Keywords: Aspose.Cells | C# | .NET | read‑only enumerator | Cell.IsNumericValue | numeric extraction | double values | OADate | convert string to numeric | iterate worksheet cells
// Common Searches: Aspose.Cells enumerate cells without changing collection | how to get all numeric values from a worksheet in Aspose.Cells | Cell.IsNumericValue example C# | convert numeric strings before iterating Aspose.Cells | read‑only Cells.GetEnumerator usage
// Developer Intent: Retrieve every numeric entry from a worksheet using a read‑only Cells enumerator.
// Use Cases: Gather numeric data for statistical analysis while preserving the original worksheet. | Export numeric values, including dates as OADate numbers, to external formats such as CSV. | Validate that a sheet contains only expected numeric entries before performing calculations or saving. | Create a summary report of all numeric cells without altering the workbook structure.
// AI Prompts: Write a C# function that returns a List<double> of all numeric values from a given Worksheet using Aspose.Cells GetEnumerator(). | Provide code that iterates cells read‑only, sums the numeric values, and treats date cells as OADate numbers. | Explain how Cell.IsNumericValue and Cell.DoubleValue work when enumerating cells with a read‑only enumerator in Aspose.Cells.

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

// Shows how to create a workbook, fill it with mixed data, optionally convert numeric strings, obtain a read‑only Cells enumerator, iterate without modifying the collection, use Cell.IsNumericValue to identify numbers, collect Cell.DoubleValue into a List<double>, and display the results.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate the worksheet with mixed data
        cells["A1"].PutValue(10);                 // integer
        cells["B1"].PutValue("Text");             // non‑numeric string
        cells["C1"].PutValue(3.14);               // double
        cells["A2"].PutValue(DateTime.Now);       // DateTime (numeric)
        cells["B2"].PutValue(true);               // boolean
        cells["C2"].PutValue("123");              // numeric string

        // Convert convertible strings to numeric values (optional)
        cells.ConvertStringToNumericValue();

        // Obtain a read‑only enumerator for the Cells collection
        IEnumerator enumerator = cells.GetEnumerator();

        // List to collect numeric values
        List<double> numericValues = new List<double>();

        // Iterate through cells without modifying the collection
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            if (cell != null && cell.IsNumericValue)
            {
                // DoubleValue returns the numeric representation (dates as OADate)
                numericValues.Add(cell.DoubleValue);
            }
        }

        // Display the collected numeric values
        Console.WriteLine("Numeric values in the worksheet:");
        foreach (double val in numericValues)
        {
            Console.WriteLine(val);
        }

        // Save the workbook (optional)
        workbook.Save("NumericValuesDemo.xlsx");
    }
}
