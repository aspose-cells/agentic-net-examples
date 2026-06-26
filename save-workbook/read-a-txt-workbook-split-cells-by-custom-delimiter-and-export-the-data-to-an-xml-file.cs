using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source TXT file
        string txtPath = "input.txt";

        // Custom delimiter used in the TXT file (e.g., pipe character)
        string customDelimiter = "|";

        // Load the TXT file with the custom delimiter
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.SeparatorString = customDelimiter; // set string delimiter
        Workbook workbook = new Workbook(txtPath, loadOptions);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine how many rows contain data in column A (index 0)
        int totalRows = sheet.Cells.MaxDataRow + 1; // +1 because rows are zero‑based

        // Split the content of column A into multiple columns using the same delimiter
        // Parameters: start row, start column, number of rows to process, load options (contains delimiter)
        sheet.Cells.TextToColumns(0, 0, totalRows, loadOptions);

        // Export the worksheet data to an XML file.
        // Assumes an XML map named "MyMap" has been defined in the workbook.
        string xmlMapName = "MyMap";
        string xmlOutputPath = "output.xml";
        workbook.ExportXml(xmlMapName, xmlOutputPath);

        Console.WriteLine("Data has been split and exported to XML successfully.");
    }
}