// Title: Generate a line chart from an Excel template using Smart Markers and WorkbookDesigner in Aspose.Cells for .NET
// Common Searches: how to fill a line chart in an Excel template with Smart Markers using Aspose.Cells C# | Aspose.Cells WorkbookDesigner process method to populate chart data from a template | create line chart from pre‑designed template workbook with Aspose.Cells .NET | C# code to load template.xlsx, process Smart Markers and save result.xlsx with chart
// Tags: populate line chart using Smart Markers Aspose.Cells | WorkbookDesigner process chart data C# | generate chart from Excel template Aspose.Cells | save workbook as XLSX after chart generation | template‑driven chart rendering Aspose.Cells .NET

using Aspose.Cells;

// The example loads an Excel template that already contains a line chart with Smart Markers, uses WorkbookDesigner to process the markers and populate the chart data, and then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Load the workbook template that already contains a line chart with Smart Markers
        Workbook workbook = new Workbook("Template.xlsx");

        // Process the Smart Markers to fill the chart data
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.Process();

        // Save the workbook with the generated line chart
        workbook.Save("Result.xlsx", SaveFormat.Xlsx);
    }
}
