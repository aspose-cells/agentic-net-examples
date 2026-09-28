// Title: Set an Excel worksheet zoom to 150% using Aspose.Cells for .NET (C#)
// AI Prompts: Create a workbook, set the first worksheet's Zoom property to 150, and save it as an .xlsx file with Aspose.Cells in C#. | Programmatically change the view scale of a worksheet to 150% before exporting the workbook using the Aspose.Cells .NET API. | Adjust the worksheet zoom level to 150 percent and persist the changes with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set worksheet zoom to 150% and save workbook | How to change Excel sheet zoom programmatically with Aspose.Cells .NET | C# example for adjusting worksheet Zoom property using Aspose.Cells | Increase Excel worksheet view scale to 150 percent using Aspose.Cells API
// Tags: Aspose.Cells worksheet zoom | C# set worksheet zoom 150% | Excel view scale Aspose.Cells .NET | adjust worksheet zoom Aspose.Cells

using Aspose.Cells;

// Creates a new workbook, sets the first worksheet's Zoom property to 150%, and saves the file as ZoomedWorkbook.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // default workbook with one worksheet

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the worksheet zoom level to 150%
        worksheet.Zoom = 150;

        // Save the workbook to a file
        workbook.Save("ZoomedWorkbook.xlsx");
    }
}
