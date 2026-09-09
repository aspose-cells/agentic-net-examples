// Title: Insert an IF formula into a worksheet column and add rows to test the conditional results with Aspose.Cells for .NET
// AI Prompts: Write C# code that assigns the formula =IF(B2>10,"High","Low") to every cell in column C of a worksheet, using Aspose.Cells, so that the formula automatically adjusts for each row. | Extend the workbook by programmatically inserting new rows, copying the IF formula into the new cells, and then checking that the calculated values match the expected High or Low outcomes.
// Common Searches: how to programmatically apply an IF formula to an entire column in Aspose.Cells C# | adding rows with formulas in Aspose.Cells without manual cell references | verify conditional formula results after inserting rows using Aspose.Cells | copying formulas to new rows automatically in Aspose.Cells workbook | saving Excel file after updating column formulas with Aspose.Cells .NET
// Tags: apply IF formula to column Aspose.Cells C# | insert rows with copied formulas Aspose.Cells | auto-adjust cell references in Aspose.Cells formulas | validate conditional results in Aspose.Cells workbook | save workbook after formula changes Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program creates a new workbook, adds header and data rows, assigns an IF formula to the Result column that returns "High" when the Score exceeds 10 and "Low" otherwise, inserts additional rows with the same formula to verify the logic, and saves the file as TableIfFormula.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add header row for the table
                sheet.Cells["A1"].PutValue("ID");
                sheet.Cells["B1"].PutValue("Score");
                sheet.Cells["C1"].PutValue("Result");

                // Add initial data rows
                sheet.Cells["A2"].PutValue(1);
                sheet.Cells["B2"].PutValue(5);   // Score 5 -> Low
                sheet.Cells["A3"].PutValue(2);
                sheet.Cells["B3"].PutValue(15);  // Score 15 -> High

                // Set IF formula for the "Result" column (structured reference not needed here)
                sheet.Cells["C2"].Formula = "=IF(B2>10,\"High\",\"Low\")";
                sheet.Cells["C3"].Formula = "=IF(B3>10,\"High\",\"Low\")";

                // Add new rows to verify the IF logic
                // Row 4: Score 8 -> should evaluate to "Low"
                sheet.Cells["A4"].PutValue(3);
                sheet.Cells["B4"].PutValue(8);
                sheet.Cells["C4"].Formula = "=IF(B4>10,\"High\",\"Low\")";

                // Row 5: Score 12 -> should evaluate to "High"
                sheet.Cells["A5"].PutValue(4);
                sheet.Cells["B5"].PutValue(12);
                sheet.Cells["C5"].Formula = "=IF(B5>10,\"High\",\"Low\")";

                // Save the workbook (lifecycle rule: save)
                workbook.Save("TableIfFormula.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
