// Title: Extract the exact formula string from an Excel cell using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells get formula string from a specific cell | How to read the raw formula of an Excel cell with Aspose.Cells .NET | Retrieve formula text of cell A1 in a workbook using Aspose.Cells C# example | Aspose.Cells Formula property vs FormulaText to obtain Excel formula as plain text | Programmatically extract Excel formula as text in C# with Aspose
// Tags: Aspose.Cells Formula property extraction C# | read raw Excel formula as text .NET | cell.Formula usage example | retrieve Excel formula string programmatically | Aspose.Cells workbook cell formula retrieval

using Aspose.Cells;
using System;

// The sample creates a workbook, assigns a SUM formula to cell A1, then uses the cell's Formula property to obtain the exact formula text and prints it to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Place a formula in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.Formula = "=SUM(B1:B5)";

            // Retrieve the formula text
            string formulaText = cell.Formula; // Use Formula property to get the textual representation

            // Display the formula text
            Console.WriteLine("Formula in A1: " + formulaText);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
