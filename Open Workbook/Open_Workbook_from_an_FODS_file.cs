using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the FODS file
        string fodsPath = "sample.fods";

        // Create load options specifying the FODS format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Fods);

        // Open the workbook from the FODS file using the constructor with file path and load options
        Workbook workbook = new Workbook(fodsPath, loadOptions);

        // Access the first worksheet and display some information
        Worksheet worksheet = workbook.Worksheets[0];
        Console.WriteLine("Worksheet Name: " + worksheet.Name);
        Console.WriteLine("Cell A1 Value: " + worksheet.Cells["A1"].StringValue);
    }
}