// Title: Freeze header columns in a worksheet and export the workbook as a macro‑enabled XLSM file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that freezes the first two columns of a worksheet and then saves the workbook as a macro‑enabled XLSM file with Aspose.Cells. | Show an example of applying column freeze to header columns and exporting the result to an .xlsm file using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# freeze first two columns and save as macro enabled workbook | how to use FreezePanes to lock header columns and export to XLSM in .NET | C# example for column freeze and XLSM export with Aspose.Cells | freeze panes header columns then save workbook as .xlsm using Aspose.Cells
// Tags: freeze panes columns Aspose.Cells | save workbook as XLSM Aspose.Cells | column freeze header Aspose.Cells C# | export frozen worksheet to macro enabled XLSM

using System;
using System.IO;
using Aspose.Cells;

// Creates a new workbook, freezes the first two columns (A and B) on the first worksheet, and saves the file as a macro‑enabled XLSM document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];

            // (Optional) Populate some data for demonstration
            // sheet.Cells["A1"].PutValue("Header1");
            // sheet.Cells["B1"].PutValue("Header2");
            // sheet.Cells["C1"].PutValue("Data1");
            // sheet.Cells["C2"].PutValue(123);

            // Freeze the first two columns (A and B) while keeping all rows unfrozen
            // Parameters: row index, column index where the freeze starts,
            // total rows to freeze, total columns to freeze.
            sheet.FreezePanes(0, 2, 0, 2);

            // Define output file path
            string outputPath = "output.xlsm";

            // Save the workbook as a macro‑enabled XLSM file
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
