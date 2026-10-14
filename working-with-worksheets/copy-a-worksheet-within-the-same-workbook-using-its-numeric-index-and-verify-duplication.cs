// Title: Copy a worksheet by its numeric index within the same workbook and verify the duplicated cells using Aspose.Cells for .NET (C#)
// AI Prompts: Use Workbook.Worksheets.AddCopy with the source sheet index to duplicate the worksheet and obtain the new sheet index. | Read specific cells from the original and the copied worksheets and compare their values to confirm the copy succeeded. | Save the workbook to a .xlsx file after performing the worksheet copy operation.
// Common Searches: Aspose.Cells C# copy worksheet by index and get new sheet index | How to check that a copied worksheet contains the same data in Aspose.Cells | AddCopy method example returning duplicated sheet index Aspose.Cells for .NET | Copy worksheet to end of workbook using numeric index Aspose.Cells C#
// Tags: addcopy worksheet index Aspose.Cells | verify worksheet copy cell values C# | duplicate worksheet within same workbook Aspose.Cells | save workbook after copying sheet Aspose.Cells | compare source and copied worksheet cells Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, adds data to the first worksheet, copies that worksheet using its numeric index with AddCopy, retrieves the new sheet, compares cell values to ensure the copy is identical, prints the verification result, and saves the workbook as CopyWorksheetDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0) and add sample data
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "SourceSheet";
            sourceSheet.Cells["A1"].PutValue("Hello");
            sourceSheet.Cells["B2"].PutValue(123);

            // Copy the worksheet to a new position (added at the end)
            // AddCopy returns the index of the newly created sheet
            int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Index);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

            // Verify duplication by comparing cell values of the original and the copy
            bool isDuplicate = sourceSheet.Cells["A1"].StringValue == copiedSheet.Cells["A1"].StringValue &&
                               sourceSheet.Cells["B2"].IntValue == copiedSheet.Cells["B2"].IntValue;

            Console.WriteLine($"Worksheet copied successfully: {isDuplicate}");

            // Save the workbook to verify the result (optional)
            workbook.Save("CopyWorksheetDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
