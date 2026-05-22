using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to cell A1 and a hyperlink pointing to an external URL
        worksheet.Cells["A1"].PutValue("Source Data");
        worksheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");

        // Configure copy options to extend the hyperlink range when copying rows
        CopyOptions copyOptions = new CopyOptions();
        copyOptions.ExtendToAdjacentRange = true;

        // Copy row 0 (the row containing the hyperlink) to row 1
        // Parameters: sourceCells, sourceRowIndex, destinationRowIndex, rowNumber, copyOptions
        worksheet.Cells.CopyRows(worksheet.Cells, 0, 1, 1, copyOptions);

        // After copying, the hyperlink count should remain the same (range extended, not duplicated)
        Console.WriteLine("Hyperlink count after copy: " + worksheet.Hyperlinks.Count);

        // Retrieve the (single) hyperlink and display its properties to confirm correct extension
        Hyperlink hyperlink = worksheet.Hyperlinks[0];
        Console.WriteLine("Hyperlink address: " + hyperlink.Address);
        Console.WriteLine("Hyperlink start row: " + hyperlink.Area.StartRow);
        Console.WriteLine("Hyperlink end row: " + hyperlink.Area.EndRow);
        Console.WriteLine("Hyperlink covers rows: " + (hyperlink.Area.EndRow - hyperlink.Area.StartRow + 1));

        // Save the workbook to verify the result manually if needed
        workbook.Save("HyperlinkCopyResult.xlsx");
    }
}