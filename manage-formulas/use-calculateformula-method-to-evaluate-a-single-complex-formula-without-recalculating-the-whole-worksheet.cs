// Title: Evaluate a complex Excel formula instantly with Worksheet.CalculateFormula in Aspose.Cells (C#)
// Description: Shows how to create a workbook, fill cells A1‑A3 and B1‑B3, define the formula =SUM(A1:A3)*AVERAGE(B1:B3), and use Worksheet.CalculateFormula to obtain the result without inserting the formula into any worksheet cell.
// Keywords: Aspose.Cells | Worksheet.CalculateFormula | C# evaluate Excel formula | calculate formula without cell | SUM and AVERAGE in Aspose.Cells | fast formula evaluation | no full recalculation
// Common Searches: Worksheet.CalculateFormula example C# | evaluate Excel formula programmatically Aspose.Cells | calculate SUM*AVERAGE without writing to cell | how to compute formula result in Aspose.Cells | C# Aspose.Cells evaluate complex formula
// Developer Intent: Compute the result of a specific Excel expression directly in code, avoiding any changes to the worksheet layout or a full workbook recalculation.
// Use Cases: Generate financial indicators from temporary data without persisting formulas. | Validate imported spreadsheet values by comparing expected and calculated results. | Run quick what‑if scenarios on ad‑hoc data sets in a high‑performance service.
// AI Prompts: Provide error‑handling code for Worksheet.CalculateFormula when the formula syntax is invalid. | Show how to use CalculateFormula with named ranges and custom user‑defined functions. | Explain how to detect the returned data type from CalculateFormula and safely convert it to double or decimal.

using System;
using Aspose.Cells;

// Shows how to create a workbook, fill cells A1‑A3 and B1‑B3, define the formula =SUM(A1:A3)*AVERAGE(B1:B3), and use Worksheet.CalculateFormula to obtain the result without inserting the formula into any worksheet cell.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some sample data that the formula will reference
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["A3"].PutValue(15);
        sheet.Cells["B1"].PutValue(2);
        sheet.Cells["B2"].PutValue(4);
        sheet.Cells["B3"].PutValue(6);

        // Define a complex formula to evaluate
        string formula = "=SUM(A1:A3)*AVERAGE(B1:B3)";

        // Evaluate the formula directly without writing it to any cell
        object result = sheet.CalculateFormula(formula);

        // Display the calculated result
        Console.WriteLine($"Result of formula \"{formula}\" is {result}");
    }
}
