// Title: Use a C# using block to automatically dispose an Aspose.Cells Workbook and release unmanaged resources
// Common Searches: C# using statement to automatically dispose Aspose.Cells Workbook | how to ensure Aspose.Cells Workbook releases file handles in .NET | best practice for disposing Aspose.Cells Workbook objects | using block for Aspose.Cells workbook resource cleanup | prevent memory leaks with Aspose.Cells Workbook in C#
// Tags: Aspose.Cells workbook disposal pattern | C# using block for Excel workbook | unmanaged resource management Aspose.Cells | auto dispose workbook in .NET | resource cleanup Aspose.Cells API

using Aspose.Cells;
using System;

// The example creates an Aspose.Cells Workbook inside a C# using block, renames the first worksheet, saves the file as output.xlsx, and relies on the using statement to automatically dispose the Workbook, ensuring unmanaged resources are released promptly.
class Program
{
    static void Main()
    {
        // Create a new workbook inside a using block to ensure it is disposed promptly
        using (var workbook = new Workbook())
        {
            // Example operation: rename the first worksheet
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        } // Workbook is automatically disposed here, releasing unmanaged resources
    }
}
