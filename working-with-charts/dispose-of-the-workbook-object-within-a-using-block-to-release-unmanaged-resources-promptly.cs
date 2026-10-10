// Title: Dispose an Aspose.Cells Workbook in C# using a using block after saving to XLSX
// AI Prompts: Generate C# code that creates an Aspose.Cells Workbook inside a using statement, writes data to a worksheet, saves the file as XLSX, and relies on the using block for automatic disposal. | Show how to wrap Aspose.Cells Workbook creation and Save operation in a using block to guarantee unmanaged resources are released immediately.
// Common Searches: how to ensure Aspose.Cells workbook is disposed after saving in C# | C# using statement for cleaning up Aspose.Cells workbook resources | best practice for releasing unmanaged resources with Aspose.Cells | automatic workbook cleanup using Aspose.Cells and using block
// Tags: using statement Aspose.Cells workbook disposal | release unmanaged resources Aspose.Cells C# | save workbook to XLSX with automatic cleanup | resource management Aspose.Cells workbook

using System;
using Aspose.Cells;

// The example creates an Aspose.Cells Workbook inside a C# using block, writes a value to cell A1, saves the workbook as output.xlsx, and the using block automatically disposes the Workbook to free unmanaged resources.
class Program
{
    static void Main()
    {
        // Create and work with the workbook inside a using block.
        // The using statement ensures that the Workbook is disposed
        // promptly, releasing unmanaged resources.
        using (var workbook = new Workbook())
        {
            // Example operation: write a value to the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells");

            // Save the workbook to a file.
            workbook.Save("output.xlsx");
        } // The workbook is automatically disposed here.
    }
}
