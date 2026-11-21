// Title: Check if Worksheet PageSetup.PrinterSettings is null with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to check if PageSetup.PrinterSettings is null | determine default printer settings value for a worksheet using Aspose.Cells | read Worksheet.PageSetup.PrinterSettings property in Aspose.Cells .NET | verify that PrinterSettings property is unassigned after workbook creation in C#
// Tags: worksheet pagesetup printersettings uninitialized | aspose.cells read printersettings property c# | default printersettings value asp.net workbook | verify printersettings not set aspose.cells

using System;
using System.Drawing.Printing;
using Aspose.Cells;

// The example creates a new Workbook, accesses its first Worksheet, and reads the Worksheet's PageSetup.PrinterSettings property to confirm that it is null (uninitialized) by default, outputting the result to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // The PrinterSettings property is not publicly settable in this version of Aspose.Cells.
            // It is null by default, so we can only read its value.
            if (sheet.PageSetup.PrinterSettings == null)
            {
                Console.WriteLine("PrinterSettings is null by default.");
            }
            else
            {
                Console.WriteLine("PrinterSettings is already assigned.");
            }

            // No further action is required; the example demonstrates that the property remains null.
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
