using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Initialize a new workbook instance
        Workbook workbook = new Workbook();

        // Import XML data from a file into the first worksheet starting at cell A1 (row 0, column 0)
        // Parameters: xml file path, destination sheet name, start row, start column
        workbook.ImportXml("data.xml", "Sheet1", 0, 0);

        // Save the workbook in XLSX format after the import operation
        workbook.Save("ImportedWorkbook.xlsx", SaveFormat.Xlsx);

        Console.WriteLine("Workbook has been saved as 'ImportedWorkbook.xlsx'.");
    }
}