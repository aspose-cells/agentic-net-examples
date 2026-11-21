// Title: Insert a scalar smart marker into a worksheet cell and populate it using WorkbookDesigner in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# place scalar smart marker '&=TotalAmount' in cell A1 | How to bind a numeric value to a smart marker using WorkbookDesigner in .NET | Replace variable marker with scalar value when processing smart markers in Aspose.Cells | Save Excel file after smart marker processing with Aspose.Cells for C# | Set data source for a single-value smart marker in Aspose.Cells workbook
// Tags: scalar smart marker insertion Aspose.Cells | WorkbookDesigner bind scalar value | smart marker processing workflow C# | populate Excel cell via variable marker | export workbook after smart marker replacement

using Aspose.Cells;

// The example creates a new workbook, writes the scalar smart marker '&=TotalAmount' into cell A1, assigns the value 1234.56 to the marker with WorkbookDesigner.SetDataSource, processes the smart markers to replace the marker with the value, and saves the result as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Place a variable marker in cell A1 to be replaced by a scalar value named "TotalAmount"
        sheet.Cells["A1"].PutValue("&=TotalAmount");

        // Initialize WorkbookDesigner to process smart markers
        var designer = new WorkbookDesigner(workbook);

        // Set the scalar data source for the marker
        designer.SetDataSource("TotalAmount", 1234.56);

        // Process the smart markers and populate the cell
        designer.Process();

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
