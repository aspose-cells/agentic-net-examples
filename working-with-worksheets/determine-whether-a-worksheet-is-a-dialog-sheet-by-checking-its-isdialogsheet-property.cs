// Title: Check if a Worksheet Is a Dialog Sheet in Aspose.Cells for .NET (IsDialogSheet Property & SheetType Fallback)
// Description: Creates a workbook, adds a dialog sheet, then demonstrates how to verify whether a worksheet is a dialog sheet using the IsDialogSheet property when available and falling back to a SheetType.Dialog comparison. The result is printed and the workbook can be saved.
// Keywords: Aspose.Cells | .NET | C# | IsDialogSheet | SheetType.Dialog | dialog sheet detection | worksheet type check | version‑compatible Aspose.Cells | worksheet properties
// Common Searches: Aspose.Cells check dialog sheet | IsDialogSheet property example | how to detect dialog sheet in Aspose.Cells | SheetType.Dialog vs IsDialogSheet | C# Aspose.Cells worksheet type
// Developer Intent: Determine whether a specific worksheet in an Aspose.Cells workbook is a dialog sheet.
// Use Cases: Validate a worksheet before applying dialog‑sheet‑specific formatting or controls. | Iterate through all worksheets and handle dialog sheets differently from regular sheets. | Provide a backward‑compatible check that uses IsDialogSheet when present and otherwise compares SheetType.Dialog.
// AI Prompts: Generate C# code that iterates through an Aspose.Cells workbook and identifies dialog sheets using IsDialogSheet with a SheetType fallback. | Show a version‑safe method to detect a dialog sheet in Aspose.Cells for .NET, handling cases where IsDialogSheet is unavailable. | Write a reusable function that returns true if a Worksheet object represents a dialog sheet, supporting both property and type checks.

using Aspose.Cells;
using System;

// Creates a workbook, adds a dialog sheet, then demonstrates how to verify whether a worksheet is a dialog sheet using the IsDialogSheet property when available and falling back to a SheetType.Dialog comparison. The result is printed and the workbook can be saved.
class CheckDialogSheet
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a dialog sheet to the workbook and obtain its index
            int dialogSheetIndex = workbook.Worksheets.Add(SheetType.Dialog);
            Worksheet dialogWorksheet = workbook.Worksheets[dialogSheetIndex];
            dialogWorksheet.Name = "MyDialogSheet";

            // Access the first worksheet (could be the default sheet or the dialog sheet)
            Worksheet worksheet = workbook.Worksheets[0];

            // Determine whether the worksheet is a dialog sheet.
            // If the IsDialogSheet property exists, it can be used directly.
            // Fallback: compare the worksheet type with SheetType.Dialog.
            bool isDialogSheet = false;

            // Uncomment the following line if the IsDialogSheet property is available in your version:
            // isDialogSheet = worksheet.IsDialogSheet;

            // Fallback check using the Type property
            isDialogSheet = worksheet.Type == SheetType.Dialog;

            Console.WriteLine($"Worksheet \"{worksheet.Name}\" is a dialog sheet: {isDialogSheet}");

            // Save the workbook (optional)
            workbook.Save("DialogSheetCheck.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
