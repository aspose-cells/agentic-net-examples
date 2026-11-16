// Title: How to set an Excel worksheet tab color to blue with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet tab color to blue programmatically | change Excel sheet tab color using Aspose.Cells .NET API | set tab color of a worksheet in a generated workbook with Aspose.Cells | C# code to apply blue tab color to a worksheet using Aspose.Cells | Aspose.Cells workbook tab color customization example
// Tags: Aspose.Cells worksheet tab color customization | apply blue tab color Excel using C# | modify worksheet tab appearance Aspose.Cells | Excel workbook tab color .NET API | colorize worksheet tabs programmatically

using Aspose.Cells;
using System.Drawing;

// Creates a new Workbook, accesses the first Worksheet, assigns Color.Blue to its TabColor property, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set the worksheet's tab color to blue
        sheet.TabColor = Color.Blue;

        // Save the workbook
        workbook.Save("Result.xlsx");
    }
}
