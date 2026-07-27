using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the default first worksheet and set its name and some data
        Worksheet originalSheet = workbook.Worksheets[0];
        originalSheet.Name = "Original";
        originalSheet.Cells["A1"].PutValue("Sample Text");
        originalSheet.Cells["B2"].PutValue(42);

        // Copy the worksheet within the same workbook using its name
        int copiedIndex = workbook.Worksheets.AddCopy("Original");
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
        copiedSheet.Name = "CopiedSheet";

        // Verify that the copied content matches the original
        bool a1Match = originalSheet.Cells["A1"].StringValue == copiedSheet.Cells["A1"].StringValue;
        bool b2Match = originalSheet.Cells["B2"].IntValue == copiedSheet.Cells["B2"].IntValue;

        Console.WriteLine($"Content integrity check: A1 match = {a1Match}, B2 match = {b2Match}");

        // Save the workbook to a file
        workbook.Save("WorksheetCopyResult.xlsx");
    }
}

// Author: Aspose.Cells .NET example