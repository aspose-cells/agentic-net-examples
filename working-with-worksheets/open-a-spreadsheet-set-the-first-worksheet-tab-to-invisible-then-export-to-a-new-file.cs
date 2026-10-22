// Title: Make the first worksheet tab invisible and export the workbook to a new .xlsx file with Aspose.Cells for .NET
// AI Prompts: Load an existing .xlsx, set the first worksheet's IsVisible property to false, and save the result to a different file using Aspose.Cells in C#. | Use Aspose.Cells to hide a specific worksheet tab before exporting the workbook to a new Excel file. | Programmatically make a worksheet invisible and create a copy of the workbook with the hidden sheet using C# and Aspose.Cells.
// Common Searches: Aspose.Cells C# hide first worksheet tab and save as new file | How to set worksheet visibility to false before exporting workbook with Aspose.Cells | Create a copy of an Excel workbook with a hidden sheet using Aspose.Cells for .NET | C# code to make a worksheet tab invisible and save workbook to another .xlsx
// Tags: worksheet visibility false Aspose.Cells | export modified workbook to new .xlsx Aspose.Cells | hide first sheet tab C# Aspose.Cells | set worksheet visibility before save .NET | Aspose.Cells copy workbook with hidden sheet

using System;
using Aspose.Cells;

// Loads input.xlsx, sets the first worksheet's IsVisible property to false to hide its tab, and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Set the first worksheet tab to invisible
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.IsVisible = false;

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
