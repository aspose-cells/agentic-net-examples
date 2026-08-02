// Title: Copy a Worksheet by Index with Aspose.Cells for .NET and Validate the Duplicate
// Description: Demonstrates how to use Aspose.Cells in C# to copy a worksheet using its numeric index (Worksheets.AddCopy), rename the copy, compare cell values to confirm identical content, modify the copy without affecting the original, and save the workbook as an Excel file.
// Keywords: Aspose.Cells C# copy worksheet by index | Worksheet.AddCopy example | duplicate sheet verification .NET | independent worksheet copy Aspose | Excel automation copy sheet | validate copied worksheet data | save workbook Aspose.Cells | C# Aspose.Cells sample code
// Common Searches: Aspose.Cells copy worksheet using index | How to verify duplicated worksheet data in .NET | C# add copy of worksheet Aspose.Cells | Ensure copied sheet is independent Aspose | Save workbook after copying sheet Aspose.Cells
// Developer Intent: Create an exact copy of an existing worksheet by its numeric index, confirm that all cell values are duplicated, and prove that changes to the copy do not impact the original sheet.
// Use Cases: Generate a template sheet once and duplicate it for each month’s report while preserving the original layout. | Create a backup of a worksheet before applying bulk transformations to prevent data loss. | Reuse formulas, styles, and charts across multiple generated reports by copying a source sheet programmatically.
// AI Prompts: Write C# code with Aspose.Cells that copies a worksheet by its numeric index, compares cell values to verify duplication, modifies the copy, and saves the workbook. | Explain step‑by‑step how Worksheets.AddCopy works, how to check that the copied sheet contains the same data, and why modifications to the copy are isolated from the original.

using System;
using Aspose.Cells;

namespace WorksheetCopyDemo
{
    // Demonstrates how to use Aspose.Cells in C# to copy a worksheet using its numeric index (Worksheets.AddCopy), rename the copy, compare cell values to confirm identical content, modify the copy without affecting the original, and save the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default contains one worksheet)
            Workbook workbook = new Workbook();

            // Access the first worksheet (index 0) and put some data
            Worksheet originalSheet = workbook.Worksheets[0];
            originalSheet.Name = "Original";
            originalSheet.Cells["A1"].PutValue("This is the original sheet");
            originalSheet.Cells["B2"].PutValue(12345);

            // Copy the worksheet using its numeric index (0)
            int copiedIndex = workbook.Worksheets.AddCopy(0);
            Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
            copiedSheet.Name = "Copied";

            // Verify that the copy contains the same data as the original
            string originalValue = originalSheet.Cells["A1"].StringValue;
            string copiedValue = copiedSheet.Cells["A1"].StringValue;
            double originalNumber = originalSheet.Cells["B2"].DoubleValue;
            double copiedNumber = copiedSheet.Cells["B2"].DoubleValue;

            Console.WriteLine($"Original A1: {originalValue}");
            Console.WriteLine($"Copied   A1: {copiedValue}");
            Console.WriteLine($"Original B2: {originalNumber}");
            Console.WriteLine($"Copied   B2: {copiedNumber}");

            // Modify the copied sheet to demonstrate it is a separate instance
            copiedSheet.Cells["A1"].PutValue("This is the copied sheet");
            copiedSheet.Cells["B2"].PutValue(98765);

            // Show that the original sheet remains unchanged
            Console.WriteLine("\nAfter modifying the copied sheet:");
            Console.WriteLine($"Original A1: {originalSheet.Cells["A1"].StringValue}");
            Console.WriteLine($"Copied   A1: {copiedSheet.Cells["A1"].StringValue}");
            Console.WriteLine($"Original B2: {originalSheet.Cells["B2"].DoubleValue}");
            Console.WriteLine($"Copied   B2: {copiedSheet.Cells["B2"].DoubleValue}");

            // Save the workbook to verify the result
            workbook.Save("WorksheetCopyResult.xlsx");
        }
    }
}
