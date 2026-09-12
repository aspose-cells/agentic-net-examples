// Title: Create a new Excel workbook, add a worksheet named "Sheet1", set it active, and save as .xlsx using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to instantiate a Workbook, add a worksheet called "Sheet1", make it the active sheet, and write the file to NewWorkbook.xlsx. | Provide a concise Aspose.Cells example that retrieves the index of a newly added worksheet and assigns it as the workbook's active sheet before saving.
// Common Searches: asp.net core create excel file with Aspose.Cells and add a specific worksheet | c# Aspose.Cells how to set newly added sheet as active sheet | save newly created workbook as xlsx using Aspose.Cells library | example code for adding a named worksheet in Aspose.Cells C#
// Tags: create workbook Aspose.Cells C# | insert worksheet with custom name Aspose.Cells | assign active sheet index Aspose.Cells | export workbook as xlsx Aspose.Cells | handle exceptions during workbook generation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample program demonstrates how to create a new Aspose.Cells Workbook in C#, add a worksheet named "Sheet1", obtain its index, set this worksheet as the active sheet, and save the workbook to "NewWorkbook.xlsx" while handling potential exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add a new worksheet named "Sheet1"
            Worksheet sheet = workbook.Worksheets.Add("Sheet1");

            // Get the index of the newly added sheet
            int sheetIndex = sheet.Index;

            // Set the newly added sheet as the active sheet (optional)
            workbook.Worksheets.ActiveSheetIndex = sheetIndex;

            // Define output file path
            string outputPath = "NewWorkbook.xlsx";

            // Save the workbook to a file (optional)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
