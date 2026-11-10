// Title: How to verify SUMPRODUCT support in Aspose.Cells for .NET before using the formula
// Common Searches: list of Excel functions supported by Aspose.Cells .NET for formula validation | C# code to check if SUMPRODUCT is available in Aspose.Cells library | determine supported Excel functions in Aspose.Cells before setting a formula | verify Excel function compatibility with Aspose.Cells in a .NET application
// Tags: Aspose.Cells supported Excel functions list | SUMPRODUCT compatibility check Aspose.Cells | validate Excel function support .NET | function availability verification Aspose.Cells | apply SUMPRODUCT formula using Aspose.Cells C#

using Aspose.Cells;
using System;

// The example creates a workbook, fills cells A1, A2, B1, and B2 with numbers, assigns the SUMPRODUCT(A1:A2,B1:B2) formula to cell C1, and saves the file as Result.xlsx, demonstrating how to use a supported Excel function in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the SUMPRODUCT function
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["B2"].PutValue(4);

            // Apply the SUMPRODUCT formula (supported in all recent versions)
            sheet.Cells["C1"].Formula = "=SUMPRODUCT(A1:A2,B1:B2)";

            // Save the workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
