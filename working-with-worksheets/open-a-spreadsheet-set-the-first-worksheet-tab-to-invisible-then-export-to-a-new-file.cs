// Title: Hide the first worksheet tab and export the workbook to a new file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells hide first worksheet tab C# and save as new workbook | set worksheet IsVisible false and export workbook using Aspose.Cells .NET | how to make a worksheet tab invisible before saving with Aspose.Cells | C# code to hide a sheet tab and create a copy of the workbook | Aspose.Cells hide worksheet tab then save to different file name
// Tags: worksheet tab visibility control Aspose.Cells | hide initial worksheet Aspose.Cells C# | export workbook with hidden sheet Aspose.Cells | save workbook to new file after visibility change Aspose.Cells | IsVisible property usage Aspose.Cells

using Aspose.Cells;

// The example loads an existing XLSX file, sets the first worksheet’s IsVisible property to false to hide its tab, and then saves the workbook as a new XLSX file.
class Program
{
    static void Main()
    {
        // Load the existing spreadsheet
        Workbook workbook = new Workbook("input.xlsx");

        // Make the first worksheet tab invisible
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.IsVisible = false;

        // Export to a new file
        workbook.Save("output.xlsx");
    }
}
