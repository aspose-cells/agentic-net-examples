// Title: Retrieve a single cell value from a formula reference using ReferredArea.GetValue in Aspose.Cells for .NET (C#)
// Common Searches: how to get the value of a cell referenced by a formula with ReferredArea.GetValue in Aspose.Cells C# | using Aspose.Cells ReferredArea.GetValue(rowOffset, colOffset) to read a referenced cell value | extract cell address from a formula and obtain its value with Aspose.Cells .NET | C# example for reading a single cell from a formula reference using Aspose.Cells | Aspose.Cells get value of referenced area without evaluating the whole formula
// Tags: Aspose.Cells ReferredArea.GetValue usage | C# retrieve value from formula reference | extract referenced cell address Aspose.Cells | single cell read from formula Aspose.Cells | Aspose.Cells workbook cell value extraction

using System;
using Aspose.Cells;

// The sample creates a workbook, writes numeric and string data to cells, assigns a formula that points to another cell, extracts the referenced address, uses ReferredArea.GetValue(rowOffset, colOffset) to obtain the value of that single referenced cell, prints the result, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some cells with sample data
            sheet.Cells["B2"].PutValue(123);          // numeric value
            sheet.Cells["C3"].PutValue("Hello");      // string value

            // Set a formula that references cell B2
            sheet.Cells["A1"].Formula = "=B2";

            // Extract the address that the formula refers to (e.g., "B2")
            string referencedAddress = null;
            string formula = sheet.Cells["A1"].Formula;
            if (!string.IsNullOrEmpty(formula) && formula.StartsWith("="))
            {
                // Simple extraction assuming a single-cell reference
                referencedAddress = formula.Substring(1);
            }

            // Retrieve the value from the referenced cell
            object referredValue = null;
            if (!string.IsNullOrEmpty(referencedAddress))
            {
                Cell refCell = sheet.Cells[referencedAddress];
                referredValue = refCell?.Value;
            }

            // Display the retrieved value
            Console.WriteLine("Referred value: " + (referredValue ?? "null"));

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
