using System;
using Aspose.Cells;
using Aspose.Cells.Properties;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Set the ScaleCrop property to true so the thumbnail displays in scaled mode
        workbook.BuiltInDocumentProperties.ScaleCrop = true;

        // Verify that the property is set
        Console.WriteLine("ScaleCrop property value: " + workbook.BuiltInDocumentProperties.ScaleCrop);

        // Save the workbook as PDF; the ScaleCrop flag will be persisted in the output file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}