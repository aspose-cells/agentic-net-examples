using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data for the SUMPRODUCT calculation
        cells["A1"].PutValue(1);
        cells["A2"].PutValue(2);
        cells["B1"].PutValue(3);
        cells["B2"].PutValue(4);

        // Set a formula that uses SUMPRODUCT
        cells["C1"].Formula = "=SUMPRODUCT(A1:A2,B1:B2)";

        // Verify whether SUMPRODUCT is supported
        // HasCustomFunction returns true if the formula contains an unsupported (custom) function
        if (cells["C1"].HasCustomFunction)
        {
            Console.WriteLine("SUMPRODUCT is NOT supported by the current Aspose.Cells version.");
        }
        else
        {
            // Since the function is supported, calculate the workbook and display the result
            workbook.CalculateFormula();
            Console.WriteLine($"SUMPRODUCT result: {cells["C1"].Value}");
        }

        // Save the workbook (optional, just to demonstrate lifecycle usage)
        workbook.Save("CheckSUMPRODUCT.xlsx");
    }
}