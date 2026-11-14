// Title: Fill a horizontal Excel range with a one‑dimensional int array using variable‑array smart markers in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# fill a row with an int[] using variable array smart marker | How to assign a one‑dimensional collection to a specific range in Aspose.Cells | Create and populate range A1:E1 from an int array with Aspose.Cells .NET | Using Range.Value to write an array to a horizontal range in C# Aspose.Cells | Variable array smart marker example for Excel workbook generation in .NET
// Tags: smart marker array population Aspose.Cells | assign int array to Excel range C# | horizontal range fill using Range.Value | create range A1:E1 Aspose.Cells | export workbook after array population .NET

using Aspose.Cells;
using System;

// The example creates a new workbook, defines an int[] of values, creates a horizontal range A1:E1 on the first worksheet, assigns the array to the range via the Range.Value property, and saves the workbook as VariableArrayMarkerExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // One‑dimensional collection to be placed into the worksheet
            int[] values = new int[] { 10, 20, 30, 40, 50 };

            // Define the target range that matches the collection size (horizontal fill)
            // This creates a range covering cells A1 through E1
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:E1");

            // Fill the range with the collection values
            range.Value = values;

            // Save the workbook
            string outputPath = "VariableArrayMarkerExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
