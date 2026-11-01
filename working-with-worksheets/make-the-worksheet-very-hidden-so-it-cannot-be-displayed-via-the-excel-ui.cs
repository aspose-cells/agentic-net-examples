// Title: How to make a worksheet very hidden (cannot be shown in Excel UI) using Aspose.Cells for .NET C#
// Common Searches: Aspose.Cells C# set worksheet to very hidden so it doesn't appear in Excel UI | Hide a worksheet permanently with Aspose.Cells .NET workbook | Make Excel worksheet invisible using Aspose.Cells IsVisible property C# | Create a very hidden sheet in Aspose.Cells and save as .xlsx | Prevent users from unhiding a worksheet with Aspose.Cells API
// Tags: Aspose.Cells very hidden worksheet | C# set worksheet IsVisible false | Aspose.Cells hide sheet from Excel UI | Workbook worksheet visibility .NET | Excel very hidden sheet Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new Workbook, renames the first worksheet to "SecretData", sets its IsVisible property to false to make the sheet very hidden (so it cannot be displayed through the Excel UI), and saves the file as VeryHiddenSheet.xlsx, with exception handling for any runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SecretData";

            // Hide the worksheet (cannot be displayed via Excel UI)
            // For older Aspose.Cells versions, use IsVisible; newer versions support Visibility,
            // but IsVisible works across all supported versions.
            sheet.IsVisible = false;

            // Save the workbook
            workbook.Save("VeryHiddenSheet.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
