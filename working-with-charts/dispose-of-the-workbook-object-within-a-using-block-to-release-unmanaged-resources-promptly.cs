// Title: Wrap an Aspose.Cells Workbook in a C# using block to ensure automatic disposal and release of unmanaged resources
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook inside a using statement, writes data to a worksheet, saves the file, and relies on the using block for disposal. | Demonstrate how to use the C# using construct so that an Aspose.Cells Workbook automatically frees its unmanaged resources after saving.
// Common Searches: C# using statement for Aspose.Cells Workbook disposal | how to release unmanaged resources after saving an Aspose.Cells workbook | ensure Aspose.Cells workbook is closed automatically in .NET | best practice for disposing Aspose.Cells Workbook objects | using block example with Aspose.Cells save to Excel
// Tags: using block Aspose.Cells Workbook | automatic disposal of Aspose.Cells workbook | release unmanaged resources Aspose.Cells | C# workbook save and dispose Aspose.Cells | resource management Aspose.Cells Workbook

using Aspose.Cells;

// Shows how to place an Aspose.Cells Workbook inside a C# using block, write a value to cell A1, save the workbook as output.xlsx, and let the using block automatically dispose the workbook, freeing unmanaged resources.
class Program
{
    static void Main()
    {
        // Create a new Workbook inside a using block to ensure it is disposed promptly
        using (var workbook = new Workbook())
        {
            // Example operation: write a value to the first worksheet
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells");

            // Save the workbook to a file
            workbook.Save("output.xlsx");
        } // The Workbook is automatically disposed here, releasing unmanaged resources
    }
}
