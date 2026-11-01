// Title: How to set an Excel worksheet zoom to 75 % using Aspose.Cells for .NET (C#)
// AI Prompts: Apply a 75 % view scale to the first worksheet with Aspose.Cells in C# and save the workbook. | Change the worksheet’s display zoom to 75 % programmatically before exporting using Aspose.Cells for .NET. | Programmatically set the Excel sheet zoom level to 75 % with Aspose.Cells and write the file as .xlsx.
// Common Searches: Aspose.Cells C# example to set worksheet zoom to 75 percent | How to change Excel worksheet view scale programmatically with Aspose.Cells .NET | C# code for adjusting worksheet display zoom before saving workbook using Aspose.Cells | Set Excel sheet zoom to 75% with Aspose.Cells for .NET tutorial | Programmatic zoom level adjustment of Excel sheet using Aspose.Cells C#
// Tags: Aspose.Cells sheet view scaling | C# adjust Excel sheet display percentage | set sheet view zoom Aspose.Cells | Excel workbook zoom configuration .NET | programmatic worksheet view scale Aspose.Cells

using Aspose.Cells;

// The sample creates (or loads) a workbook, accesses the first worksheet, sets its view zoom to 75 %, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the worksheet zoom to 75 percent
        worksheet.Zoom = 75;

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
