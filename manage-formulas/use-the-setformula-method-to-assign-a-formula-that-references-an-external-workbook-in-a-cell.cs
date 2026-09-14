// Title: Use Aspose.Cells for .NET (C#) to set a cell formula that links to an external workbook
// AI Prompts: Generate C# code with Aspose.Cells that writes the formula ='[External.xlsx]Sheet1'!A1 into cell B2 and saves the workbook. | Show how to assign an external workbook reference to a cell using the Cell.Formula property in Aspose.Cells for .NET.
// Common Searches: how to reference a cell from another Excel file using Aspose.Cells C# | Aspose.Cells C# set formula to external workbook example | link external workbook cell in Aspose.Cells .NET | C# Aspose.Cells formula with external file path
// Tags: set external workbook formula Aspose.Cells | Cell.Formula property external reference | Aspose.Cells link to another Excel file | C# assign formula with external workbook

using System;
using System.IO;
using Aspose.Cells;

// // This program creates a new workbook, accesses cell B2 on the first worksheet, assigns a formula that references cell A1 in Sheet1 of an external file named External.xlsx, and saves the workbook as Result.xlsx.
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

            // Get the target cell
            Cell targetCell = sheet.Cells["B2"];

            // Assign a formula that references an external workbook (External.xlsx, Sheet1!A1)
            // Using the Formula property (SetFormula overload with one argument is not available)
            targetCell.Formula = "='[External.xlsx]Sheet1'!A1";

            // Save the workbook
            string resultPath = "Result.xlsx";
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(resultPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
