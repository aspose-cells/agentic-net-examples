// Title: How to explicitly dispose an Aspose.Cells Workbook in C# to free file handles and memory
// Common Searches: how to call Workbook.Dispose after saving an Excel file with Aspose.Cells in C# | C# Aspose.Cells release file lock after workbook.Save | prevent memory usage growth when generating multiple workbooks with Aspose.Cells | best practice for disposing Aspose.Cells Workbook objects in .NET applications | why does Excel file remain locked after using Aspose.Cells Save method
// Tags: Aspose.Cells workbook disposal pattern | prevent memory leaks Aspose.Cells .NET | explicit Workbook.Dispose usage | clean up Aspose.Cells resources | manage Excel file handles with Aspose.Cells

using System;
using Aspose.Cells;

// The example creates an Aspose.Cells Workbook, fills a worksheet with sample data, saves it as Output.xlsx, and then calls workbook.Dispose() to explicitly release file handles and memory, ensuring no lingering locks or memory leaks.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "SampleData";

        // Add some data to cells
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");

        // Explicitly release resources
        workbook.Dispose();
    }
}
