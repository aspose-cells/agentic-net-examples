// Title: Hide an Excel worksheet in a .NET workbook with Aspose.Cells while keeping it usable by VBA macros
// Common Searches: how to hide a worksheet using Aspose.Cells C# but still allow VBA macros to access it | Aspose.Cells set worksheet visibility to hidden for .xlsx files in C# | C# code example for normal hidden Excel sheet with Aspose.Cells that VBA can read | saving a hidden worksheet with Aspose.Cells .NET and preserving macro access | Aspose.Cells hide worksheet without using VeryHidden in C#
// Tags: Aspose.Cells hide worksheet | Aspose.Cells normal hidden sheet | Aspose.Cells worksheet visibility C# | Aspose.Cells VBA accessible hidden sheet | Aspose.Cells save hidden worksheet to xlsx

using Aspose.Cells;
using System;
using System.IO;

// The example creates a new Workbook, adds data to the first worksheet, sets its IsVisible property to false (normal hidden), and saves the file as HiddenSheet.xlsx. The hidden sheet remains accessible to VBA macros while being invisible to end users.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and give it a name
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add some sample data to the worksheet
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");

            // Hide the worksheet (normal hidden; VeryHidden not available in this version)
            sheet.IsVisible = false;

            // Define output file path
            string outputPath = "HiddenSheet.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
