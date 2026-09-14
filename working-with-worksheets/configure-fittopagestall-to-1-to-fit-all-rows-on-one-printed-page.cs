// Title: Set FitToPagesTall = 1 to print all worksheet rows on a single page with Aspose.Cells for .NET
// AI Prompts: Configure a worksheet's PageSetup.FitToPagesTall property to 1 using Aspose.Cells in C# and save the workbook. | Create an Excel file where the first sheet prints all rows on one page by setting FitToPagesTall to 1 with the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# set FitToPagesTall to 1 for single page printing | how to fit all rows on one printed page using Aspose.Cells .NET | PageSetup FitToPagesTall example Aspose.Cells C# | print Excel worksheet with all rows on one page Aspose.Cells
// Tags: Aspose.Cells worksheet FitToPagesTall setting | C# page setup fit rows to one page | Aspose.Cells single-page print layout | FitToPagesTall property usage Aspose.Cells | generate Excel file with one-page row fit C#

using Aspose.Cells;
using System;

// The example creates a new workbook, sets the first worksheet’s PageSetup.FitToPagesTall to 1 so all rows print on a single page, and saves the file as FitToPagesTall.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Configure FitToPagesTall to 1 so all rows fit on a single printed page
        sheet.PageSetup.FitToPagesTall = 1;

        // Save the workbook
        workbook.Save("FitToPagesTall.xlsx", SaveFormat.Xlsx);
    }
}
