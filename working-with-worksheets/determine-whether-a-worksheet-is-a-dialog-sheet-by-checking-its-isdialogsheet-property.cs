// Title: How to determine if a worksheet is a dialog sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells and returns true if the first worksheet’s Type equals SheetType.Dialog. | Show how to use the Worksheet.Type property and SheetType enumeration to verify a dialog sheet in Aspose.Cells. | Create a console program that prints the dialog‑sheet status of a specified worksheet in an .xlsx file using Aspose.Cells.
// Common Searches: aspocells c# check if worksheet is a dialog sheet | how to use SheetType.Dialog to identify dialog sheets in Excel files | determine worksheet type dialog Aspose.Cells .NET example | C# program to read Excel and detect dialog sheet with Aspose | isdialogsheet property Aspose.Cells usage tutorial
// Tags: dialog sheet detection Aspose.Cells | worksheet type check SheetType.Dialog C# | use Worksheet.Type property Aspose.Cells | identify dialog sheet in .xlsx with Aspose | console app worksheet dialog status Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook with Aspose.Cells, accesses a worksheet, checks if its Type equals SheetType.Dialog, and outputs the result to the console.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (by index or name)
            Worksheet sheet = workbook.Worksheets[0]; // or workbook.Worksheets["Sheet1"]

            // Determine if the worksheet is a dialog sheet using the SheetType enumeration
            bool isDialogSheet = sheet.Type == SheetType.Dialog;

            // Output the result
            Console.WriteLine($"Worksheet \"{sheet.Name}\" IsDialogSheet: {isDialogSheet}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
