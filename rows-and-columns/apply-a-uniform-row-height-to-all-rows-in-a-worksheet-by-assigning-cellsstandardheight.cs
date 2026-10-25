// Title: How to set a uniform row height for every row in an Aspose.Cells worksheet using Cells.StandardHeight in C#
// Common Searches: Aspose.Cells C# set same row height for all rows in a worksheet | using Cells.StandardHeight to define default row height in .NET Excel file | apply global row height to Excel worksheet with Aspose.Cells API | C# code to set uniform row height across entire sheet using Aspose.Cells | change default row height for new rows in Aspose.Cells workbook
// Tags: uniform row height Aspose.Cells C# | Cells.StandardHeight usage .NET | global worksheet row height Aspose.Cells | default Excel row height C# Aspose | apply row height to entire sheet Aspose.Cells

using Aspose.Cells;

// The example creates a new Workbook, accesses its first Worksheet, sets the Cells.StandardHeight property to 20 points, which makes every row in the sheet have the same height, and then saves the file as UniformRowHeight.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Apply a uniform row height (e.g., 20 points) to all rows in the worksheet
        worksheet.Cells.StandardHeight = 20;

        // Save the workbook to a file
        workbook.Save("UniformRowHeight.xlsx");
    }
}
