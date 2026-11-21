// Title: Use a C# using statement to load and save an Excel workbook with Aspose.Cells for deterministic disposal
// Common Searches: how to use a using block with Aspose.Cells Workbook in C# | ensure workbook object is disposed after saving Excel file using Aspose.Cells | C# example of loading and saving Excel with deterministic resource cleanup | Aspose.Cells load workbook inside using statement for .NET | best practice for disposing Aspose.Cells Workbook objects
// Tags: using statement Aspose.Cells workbook disposal | load and save Excel with Aspose.Cells C# | deterministic resource cleanup Aspose.Cells | Aspose.Cells workbook using block .NET | Excel file handling Aspose.Cells C#

using System;
using Aspose.Cells;

// Demonstrates loading an Excel file into an Aspose.Cells Workbook inside a C# using block, performing optional worksheet operations, saving to a new file, and relying on the using statement for automatic, deterministic disposal of the workbook resource.
class Program
{
    static void Main()
    {
        // Load the workbook inside a using block to guarantee deterministic disposal
        using (var workbook = new Workbook("input.xlsx"))
        {
            // Example operation: access the first worksheet
            var worksheet = workbook.Worksheets[0];
            // TODO: add custom manipulation of the worksheet here

            // Save the workbook; the same using block ensures the object is disposed after saving
            workbook.Save("output.xlsx");
        } // workbook is automatically disposed here
    }
}
