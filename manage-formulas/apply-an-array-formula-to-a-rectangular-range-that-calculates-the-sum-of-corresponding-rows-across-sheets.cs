// Title: Set an array formula to sum matching rows across worksheets with Aspose.Cells for .NET
// Description: Creates a workbook, adds Sheet1, Sheet2 and Result sheets, fills A1:A5 on the first two sheets, defines the array formula "=Sheet1!A1:A5+Sheet2!A1:A5", applies it to Result!A1 using SetArrayFormula(5,1), forces calculation, and saves the workbook.
// Keywords: Aspose.Cells | SetArrayFormula | array formula | cross‑sheet sum | .NET | Excel automation | calculate formulas | workbook save
// Common Searches: Aspose.Cells SetArrayFormula across multiple sheets | how to add ranges from different worksheets using Aspose.Cells | calculate array formulas after setting them in .NET | sum corresponding rows from two sheets with Aspose.Cells
// Developer Intent: Create a rectangular array formula that adds each row from two source worksheets and writes the summed values to a third worksheet.
// Use Cases: Combine monthly sales numbers from regional sheets into a consolidated summary with a single formula. | Produce a unified financial total by adding matching rows from separate department worksheets. | Generate a dynamic report that aggregates sensor readings stored on multiple sheets without writing individual cell formulas.
// AI Prompts: Show how to use SetArrayFormula to multiply corresponding ranges from three worksheets in Aspose.Cells for .NET. | Provide code that reads the results of an array formula after calling CalculateFormula with Aspose.Cells. | Explain how to adjust SetArrayFormula for a multi‑column range that adds rows from several sheets.

using System;
using Aspose.Cells;

// Creates a workbook, adds Sheet1, Sheet2 and Result sheets, fills A1:A5 on the first two sheets, defines the array formula "=Sheet1!A1:A5+Sheet2!A1:A5", applies it to Result!A1 using SetArrayFormula(5,1), forces calculation, and saves the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();

        // Rename the default sheet and add two more sheets
        Worksheet sheet1 = workbook.Worksheets[0];
        sheet1.Name = "Sheet1";
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
        Worksheet resultSheet = workbook.Worksheets.Add("Result");

        // Populate sample data in Sheet1 (1,2,3,4,5) and Sheet2 (10,20,30,40,50)
        for (int i = 0; i < 5; i++)
        {
            sheet1.Cells[i, 0].PutValue(i + 1);          // A1:A5 in Sheet1
            sheet2.Cells[i, 0].PutValue((i + 1) * 10);   // A1:A5 in Sheet2
        }

        // Define an array formula that sums corresponding rows from Sheet1 and Sheet2
        // The formula will spill into 5 rows and 1 column on the Result sheet
        string arrayFormula = "=Sheet1!A1:A5+Sheet2!A1:A5";

        // Apply the array formula starting at Result!A1
        // Parameters: (formula, number of rows to populate, number of columns to populate)
        resultSheet.Cells["A1"].SetArrayFormula(arrayFormula, 5, 1);

        // Calculate all formulas so the array results are materialized
        workbook.CalculateFormula();

        // Save the workbook (lifecycle: save)
        workbook.Save("ArrayFormulaSumAcrossSheets.xlsx");
    }
}
