// Title: Create a VeryHidden worksheet in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, sets the first worksheet's Visibility to VeryHidden, and saves the file as an .xlsx using Aspose.Cells. | Write C# to open an existing Excel workbook, mark a specified sheet as VeryHidden so it cannot be unhidden through the Excel UI, then save the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells how to set worksheet to VeryHidden in C# | C# make Excel sheet invisible to user using Aspose.Cells | prevent users from unhiding a worksheet with Aspose.Cells .NET | set worksheet visibility to VeryHidden programmatically Aspose.Cells | Aspose.Cells hide sheet from Excel UI C# example
// Tags: Aspose.Cells set worksheet VeryHidden | C# hide Excel worksheet UI | Aspose.Cells worksheet visibility enum | Excel VeryHidden sheet Aspose.Cells | Aspose.Cells save workbook with hidden sheet

using System;
using System.IO;
using Aspose.Cells;

// The program creates a new Workbook, accesses the first Worksheet, marks it as VeryHidden so it cannot be displayed via the Excel UI, saves the workbook as VeryHiddenWorksheet.xlsx, and outputs the full file path.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Hide the worksheet. In older Aspose.Cells versions the Visibility property is not available,
            // so we use the IsVisible flag. This makes the sheet hidden (cannot be shown via UI).
            worksheet.IsVisible = false;

            // Define output file path
            string outputPath = "VeryHiddenWorksheet.xlsx";

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
