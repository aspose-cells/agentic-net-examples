// Title: Hide the first worksheet tab and export the workbook to a new Excel file with Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx file, set the first worksheet's IsVisible property to false, and save the workbook under a different filename using Aspose.Cells in C#. | Using Aspose.Cells for .NET, programmatically hide the first sheet tab and write the modified workbook to a new Excel document.
// Common Searches: Aspose.Cells C# hide first worksheet tab and save as new file | how to set worksheet visibility to hidden with Aspose.Cells .NET | export workbook after changing sheet visibility using Aspose.Cells | C# code to hide a sheet tab and create a copy of the Excel workbook
// Tags: Aspose.Cells worksheet visibility control | C# modify worksheet IsVisible attribute | save workbook under new filename Aspose.Cells | export workbook after sheet visibility change

using System;
using Aspose.Cells;

// The example loads 'input.xlsx', hides the first worksheet tab by setting its IsVisible property to false, and saves the updated workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Set the first worksheet tab to invisible
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.IsVisible = false;

        // Export (save) to a new file
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
