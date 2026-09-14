// Title: Set the worksheet top margin to 5 points using Aspose.Cells in C#
// AI Prompts: Load an Excel workbook, assign PageSetup.TopMargin = 5 (points), and save the file with Aspose.Cells in C#. | Programmatically change the top margin of a worksheet to exactly five points via the Aspose.Cells PageSetup API. | Write a .NET script that opens a .xlsx file, updates the top margin to 5 points, and writes the modified workbook.
// Common Searches: Aspose.Cells C# set worksheet top margin to 5 points | How to change Excel page setup top margin in points using Aspose.Cells .NET | C# example for adjusting worksheet top margin with Aspose.Cells | Set top margin measurement in points for Excel file via Aspose.Cells API
// Tags: Aspose.Cells page setup top margin points | C# adjust worksheet top margin | Excel top margin modification using Aspose.Cells | set worksheet top margin .NET

using Aspose.Cells;

// Loads an existing workbook, sets the first worksheet's top margin to 5 points via PageSetup.TopMargin, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Adjust the top margin to 5 points for consistent layout
        sheet.PageSetup.TopMargin = 5; // TopMargin is measured in points

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
