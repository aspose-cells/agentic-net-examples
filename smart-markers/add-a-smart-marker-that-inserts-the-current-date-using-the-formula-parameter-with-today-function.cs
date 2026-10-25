// Title: Insert current date with a smart marker using the TODAY() formula in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# smart marker to display today's date with TODAY() | how to use the Formula parameter in a smart marker for dynamic date in .NET | example of inserting a TODAY() formula via smart markers in an Excel file using Aspose.Cells | C# code to add a smart marker that writes the current date into a cell with Aspose.Cells | generate Excel workbook with smart marker date field using Aspose.Cells for .NET
// Tags: smart marker insert date Aspose.Cells | TODAY() formula smart marker .NET | Aspose.Cells generate Excel with dynamic date | C# smart marker formula parameter usage | Excel workbook creation Aspose.Cells C#

using System;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, assigns the =TODAY() formula to cell A1 via a smart marker, saves the file as SmartMarkerDate.xlsx, and prints the output path while handling any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Access the first worksheet
            var sheet = workbook.Worksheets[0];

            // Insert the TODAY() formula directly into cell A1
            sheet.Cells["A1"].Formula = "=TODAY()";

            // Save the workbook to a file
            string outputPath = "SmartMarkerDate.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
