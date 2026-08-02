// Title: Set worksheet print order to OverThenDown (portrait) using Aspose.Cells for .NET
// Description: Creates a new Workbook, sets the first worksheet to portrait orientation, changes the print order to OverThenDown via PageSetup.Order, and saves the file as an XLSX document.
// Keywords: Aspose.Cells | C# | SetPageOrder | PrintOrderType | OverThenDown | portrait orientation | page setup | Excel pagination | Workbook.Save | XLSX export | worksheet printing
// Common Searches: Aspose.Cells set print order OverThenDown | C# page setup portrait orientation Aspose.Cells | how to change worksheet pagination order in .NET | save workbook after modifying page order Aspose | print order OverThenDown example C#
// Developer Intent: Configure a worksheet to print pages over‑then‑down in portrait mode and persist the changes to an Excel file.
// Use Cases: Generating multi‑page reports where each row of pages is filled horizontally before moving to the next row. | Creating printable invoices or forms that require booklet‑style pagination. | Preparing portrait‑oriented worksheets for accurate physical page sequencing during batch printing.
// AI Prompts: Write C# code with Aspose.Cells that sets a worksheet to portrait orientation and OverThenDown print order, then saves it as XLSX. | Explain the effect of PrintOrderType.OverThenDown on Excel pagination when using Aspose.Cells. | Provide step‑by‑step instructions to configure page setup for portrait mode and over‑then‑down print order in Aspose.Cells for .NET.

using Aspose.Cells;
using System;
using System.IO;

// Creates a new Workbook, sets the first worksheet to portrait orientation, changes the print order to OverThenDown via PageSetup.Order, and saves the file as an XLSX document.
public class SetPageOrderDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the sheet is in portrait orientation (optional but typical for pagination)
            sheet.PageSetup.Orientation = PageOrientationType.Portrait;

            // Set the print order to OverThenDown for proper pagination
            sheet.PageSetup.Order = PrintOrderType.OverThenDown;

            // Define output file path
            string outputPath = "PageOrder_OverThenDown.xlsx";

            // Save the workbook to an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SetPageOrderDemo.Run();
    }
}
