// Title: Copy a Row with Hyperlink and Preserve Its Range Using Aspose.Cells for .NET
// Description: Demonstrates how to add a hyperlink to a cell, copy the entire row with Cells.CopyRows, enable ExtendToAdjacentRange, and confirm that the hyperlink area expands to include the new row before saving the workbook.
// Keywords: Aspose.Cells copy rows | hyperlink preservation .NET | ExtendToAdjacentRange | Cells.CopyRows C# | Excel hyperlink area | verify hyperlink after copy | Aspose.Cells API | C# Excel automation
// Common Searches: Aspose.Cells copy row with hyperlink | ExtendToAdjacentRange effect on hyperlinks | C# verify hyperlink range after row copy | How to keep hyperlinks when copying rows in Aspose.Cells | CopyRows preserve hyperlink area
// Developer Intent: Copy a row that contains a hyperlink and ensure the hyperlink automatically extends to the newly copied row.
// Use Cases: Duplicate a data row that includes an external link while keeping the link target unchanged. | Programmatically copy multiple rows with embedded hyperlinks and validate that the hyperlink count remains constant. | Generate a report where linked rows are moved to a different section and the hyperlink range must cover the new rows.
// AI Prompts: Provide C# code that copies rows containing hyperlinks with Aspose.Cells and keeps the link addresses intact. | Explain how CopyOptions.ExtendToAdjacentRange updates Hyperlink.Area when rows are duplicated. | Show a method to check that Hyperlink.Area.StartRow and EndRow reflect the added rows after using Cells.CopyRows.

using System;
using Aspose.Cells;

namespace AsposeCellsHyperlinkCopyDemo
{
    // Demonstrates how to add a hyperlink to a cell, copy the entire row with Cells.CopyRows, enable ExtendToAdjacentRange, and confirm that the hyperlink area expands to include the new row before saving the workbook.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data in row 2 (index 1) and a hyperlink in cell A2
            sheet.Cells["A2"].PutValue("Original Row");
            // Hyperlink points to an external website
            sheet.Hyperlinks.Add("A2", 1, 1, "https://www.example.com");

            // Display original hyperlink information
            Console.WriteLine("Before copy:");
            foreach (Hyperlink link in sheet.Hyperlinks)
            {
                Console.WriteLine($"Address: {link.Address}");
                Console.WriteLine($"Area: Row {link.Area.StartRow} - {link.Area.EndRow}, Column {link.Area.StartColumn} - {link.Area.EndColumn}");
            }

            // Prepare copy options to extend hyperlink range when copying to adjacent rows
            CopyOptions copyOptions = new CopyOptions();
            copyOptions.ExtendToAdjacentRange = true;

            // Copy row 2 (index 1) to row 3 (index 2)
            // Parameters: sourceCells, sourceRowIndex, destinationRowIndex, rowNumber, copyOptions
            sheet.Cells.CopyRows(sheet.Cells, 1, 2, 1, copyOptions);

            // Verify hyperlink count (should remain the same) and that the hyperlink area now includes the new row
            Console.WriteLine("\nAfter copy:");
            Console.WriteLine($"Hyperlink count: {sheet.Hyperlinks.Count}");
            foreach (Hyperlink link in sheet.Hyperlinks)
            {
                Console.WriteLine($"Address: {link.Address}");
                Console.WriteLine($"Area: Row {link.Area.StartRow} - {link.Area.EndRow}, Column {link.Area.StartColumn} - {link.Area.EndColumn}");
            }

            // Save the workbook to verify manually if needed
            workbook.Save("HyperlinkCopyResult.xlsx");
        }
    }
}
