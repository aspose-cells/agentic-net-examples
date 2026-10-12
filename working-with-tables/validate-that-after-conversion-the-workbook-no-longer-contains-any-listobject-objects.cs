// Title: Check for remaining ListObject tables in a workbook after converting it with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, saves it as PDF, then iterates all worksheets to detect any ListObject tables and returns a boolean result. | Create a method using Aspose.Cells for .NET that validates a workbook contains zero ListObject objects after any format conversion. | Generate a console application that converts an .xlsx file to PDF and prints a message indicating whether any worksheet still has ListObjects.
// Common Searches: how to programmatically verify that an Excel workbook has no tables after saving with Aspose.Cells | Aspose.Cells C# detect ListObject objects after exporting to PDF | C# code to ensure all ListObjects are removed from worksheets after conversion | validate workbook contains zero ListObjects using Aspose.Cells .NET API
// Tags: Aspose.Cells listobject detection after conversion | C# validate workbook tables post PDF export | check worksheet ListObjects with Aspose.Cells | ensure no Excel tables after format conversion .NET | listobject presence verification in Aspose.Cells

using System;
using Aspose.Cells;

// Loads an Excel workbook, saves it as PDF using Aspose.Cells, then scans every worksheet for ListObject tables and reports whether any remain.
class Program
{
    static void Main()
    {
        // Load the source workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Example conversion: save the workbook to PDF (any other format can be used)
        workbook.Save("output.pdf", SaveFormat.Pdf);

        // Validate that no ListObject (table) exists in any worksheet after conversion
        bool containsListObjects = false;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.ListObjects.Count > 0)
            {
                containsListObjects = true;
                break;
            }
        }

        // Output the validation result
        Console.WriteLine(containsListObjects
            ? "Validation failed: ListObject objects are still present."
            : "Validation succeeded: No ListObject objects found.");
    }
}
