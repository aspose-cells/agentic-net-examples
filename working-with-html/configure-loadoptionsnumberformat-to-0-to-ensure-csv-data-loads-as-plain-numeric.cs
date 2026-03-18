using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a sample CSV file
        string csvPath = "sample.csv";
        File.WriteAllText(csvPath, "Header1,Header2,Header3\n1,2,3\n4,5,6");

        // Create load options and set NumberFormat to "0"
        // This forces numeric values to be loaded without any formatting
        XmlLoadOptions loadOptions = new XmlLoadOptions();
        loadOptions.NumberFormat = "0";

        // Load the CSV file using the specified load options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Save the workbook as an XLSX file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}