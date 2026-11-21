// Title: Determine if an Excel worksheet is a dialog sheet (IsDialogSheet) with Aspose.Cells for .NET in C#
// Common Searches: C# Aspose.Cells how to check if a worksheet is a dialog sheet | using Aspose.Cells SheetType.Dialog to identify dialog sheets in Excel | Aspose.Cells .NET example for detecting dialog sheet type | code sample to verify worksheet IsDialogSheet property with Aspose.Cells | load workbook and test for dialog sheet using Aspose.Cells C#
// Tags: Aspose.Cells dialog sheet detection C# | SheetType.Dialog usage Aspose.Cells | check worksheet type Aspose.Cells .NET | identify dialog sheet Excel Aspose.Cells | load workbook inspect sheet type C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, accesses a worksheet, and determines whether it is a dialog sheet by comparing its Type to SheetType.Dialog (equivalent to the IsDialogSheet property). The result is printed to the console, with error handling for missing files or other exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or any worksheet you need to check)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine whether the worksheet is a dialog sheet using its SheetType
            bool isDialogSheet = sheet.Type == SheetType.Dialog;

            // Output the result
            Console.WriteLine($"Worksheet \"{sheet.Name}\" IsDialogSheet: {isDialogSheet}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
