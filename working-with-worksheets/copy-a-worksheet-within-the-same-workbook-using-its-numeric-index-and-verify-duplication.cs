// Title: Copy a worksheet by its numeric index within the same workbook and verify the duplicated data using Aspose.Cells for .NET
// AI Prompts: Invoke Workbook.Worksheets.AddCopy with the source sheet index to create a duplicate and capture the returned worksheet index. | Read cells A1 and B2 from both the original and the newly copied worksheets and compare their values to confirm the copy is identical.
// Common Searches: Aspose.Cells C# copy worksheet by index example | how to verify copied worksheet data with Aspose.Cells .NET | Get index of new worksheet after AddCopy in Aspose.Cells | duplicate sheet inside same workbook and compare cells Aspose.Cells | C# Aspose.Cells copy worksheet and check cell values
// Tags: Aspose.Cells AddCopy worksheet duplication | validate worksheet copy Aspose.Cells C# | duplicate sheet by index Aspose.Cells .NET | save workbook after copying sheet Aspose.Cells | worksheet copy integrity check Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, adds sample data to the first worksheet, duplicates that worksheet using the AddCopy method with its numeric index, retrieves the copied sheet, validates the duplication by comparing the values in cells A1 and B2 between the source and copy, and finally saves the workbook to an XLSX file.
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

            // Copy the worksheet within the same workbook
            int sourceIndex = 0;
            // AddCopy returns the index of the newly created worksheet
            int copiedIndex = workbook.Worksheets.AddCopy(sourceIndex);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];

            // Verify duplication by comparing cell values
            bool isDuplicate = sourceSheet.Cells["A1"].StringValue == copiedSheet.Cells["A1"].StringValue &&
                               sourceSheet.Cells["B2"].IntValue == copiedSheet.Cells["B2"].IntValue;

            Console.WriteLine($"Worksheet copied. Verification result: {isDuplicate}");

            // Save the workbook (optional)
            string outputPath = "CopyWorksheetDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
