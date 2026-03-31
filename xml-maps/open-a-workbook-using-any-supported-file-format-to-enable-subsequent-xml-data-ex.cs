using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to an existing workbook file (XLSX, XLS, CSV, etc.).
        // Replace with the actual file you want to open.
        string inputFile = "sample.xlsx";

        // Open the workbook using the string constructor (load rule).
        Workbook workbook = new Workbook(inputFile);

        // The workbook is now loaded and ready for XML export operations.
        // Example of exporting XML data linked to an XML map named "MyMap".
        // Uncomment and set the correct map name and output path as needed.
        // workbook.ExportXml("MyMap", "exported.xml");

        // Optional: save the workbook in another format if required.
        // workbook.Save("copy_of_workbook.xlsx", SaveFormat.Xlsx);
    }
}