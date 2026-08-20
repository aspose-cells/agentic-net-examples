// Title: Determine if a worksheet is a dialog sheet using Aspose.Cells IsDialogSheet in C#
// Description: Creates a workbook, adds a dialog sheet (SheetType.Dialog), and uses the IsDialogSheet property (or SheetType comparison) to identify dialog worksheets. The result is written to the console and the file can be saved.
// Keywords: Aspose.Cells | C# | .NET | IsDialogSheet | dialog sheet detection | SheetType.Dialog | worksheet type check | identify dialog worksheet | Aspose.Cells example
// Common Searches: Aspose.Cells IsDialogSheet C# example | how to detect dialog sheet in Aspose.Cells | check worksheet type Aspose.Cells .NET | identify dialog worksheets using Aspose.Cells | IsDialogSheet property usage
// Developer Intent: Find out whether a given worksheet is a dialog sheet.
// Use Cases: Skip dialog sheets while exporting data from a workbook. | Apply custom formatting only to dialog worksheets. | Validate workbook structure by confirming the presence of dialog sheets before publishing.
// AI Prompts: Write C# code that iterates through all worksheets in an Aspose.Cells workbook and prints the worksheet name with a true/false IsDialogSheet flag. | Show how to exclude dialog sheets when copying data between two Aspose.Cells workbooks in .NET. | Provide an example that logs the names of every dialog sheet in a workbook using Aspose.Cells for C#.

using System;
using Aspose.Cells;

// Creates a workbook, adds a dialog sheet (SheetType.Dialog), and uses the IsDialogSheet property (or SheetType comparison) to identify dialog worksheets. The result is written to the console and the file can be saved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a dialog sheet to the workbook (Add returns the sheet index)
            int dialogSheetIndex = workbook.Worksheets.Add(SheetType.Dialog);
            Worksheet dialogWorksheet = workbook.Worksheets[dialogSheetIndex];
            dialogWorksheet.Name = "MyDialogSheet";

            // The workbook already contains a default worksheet (type Worksheet)
            Worksheet normalWorksheet = workbook.Worksheets[0];

            // Check whether each worksheet is a dialog sheet using the Type property
            Console.WriteLine($"{dialogWorksheet.Name} IsDialogSheet: {dialogWorksheet.Type == SheetType.Dialog}");
            Console.WriteLine($"{normalWorksheet.Name} IsDialogSheet: {normalWorksheet.Type == SheetType.Dialog}");

            // Save the workbook (optional)
            workbook.Save("DialogSheetCheck.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
