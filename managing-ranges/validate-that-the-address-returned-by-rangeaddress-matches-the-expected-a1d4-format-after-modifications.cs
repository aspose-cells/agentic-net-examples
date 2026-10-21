// Title: Validate that a Range.Address returns the exact "A1:D4" string after modifying cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to check if Range.Address equals A1:D4 after cell updates | Verify range address format in a .NET workbook with Aspose.Cells | C# code to compare Range.Address property to expected A1:D4 string | Get and validate Excel range address using Aspose.Cells API | Confirm that a created range still reports A1:D4 after modifying its cells
// Tags: Aspose.Cells range address validation | C# Aspose.Cells compare Range.Address | Aspose.Cells verify A1:D4 range | Aspose.Cells modify cells range address check | Aspose.Cells .NET workbook range address property

using Aspose.Cells;
using System;

// The example creates a workbook, fills cells A1 through D4, defines a range covering those cells, updates values inside the range, retrieves the range's Address property, compares it to the expected "A1:D4" string, prints the validation result, and saves the workbook as ModifiedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (using the provided creation rule)
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Populate some data within the range A1:D4
        sheet.Cells["A1"].PutValue(1);
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["C3"].PutValue(3);
        sheet.Cells["D4"].PutValue(4);

        // Create a range that covers cells A1 through D4
        var range = sheet.Cells.CreateRange("A1", "D4");

        // Modify values inside the range (the range itself is not moved)
        range[0, 0].PutValue(10); // A1
        range[3, 3].PutValue(40); // D4

        // Retrieve the address of the range after modifications
        string address = range.Address;

        // Validate that the address matches the expected A1:D4 format
        if (address == "A1:D4")
        {
            Console.WriteLine("Address validation passed: " + address);
        }
        else
        {
            Console.WriteLine($"Address validation failed. Expected A1:D4 but got {address}");
        }

        // Save the workbook (using the provided save rule)
        workbook.Save("ModifiedWorkbook.xlsx");
    }
}
