using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

class XmlMapImportParallel
{
    static void Main()
    {
        // Define the XML files to import and the corresponding output workbooks.
        var jobs = new List<(string XmlPath, string OutputPath, string SheetName)>
        {
            ("data1.xml", "book1.xlsx", "Sheet1"),
            ("data2.xml", "book2.xlsx", "Sheet1"),
            ("data3.xml", "book3.xlsx", "Sheet1")
        };

        // Process each workbook concurrently.
        Parallel.ForEach(jobs, job =>
        {
            // Create a new workbook instance.
            Workbook wb = new Workbook();

            // Enable multi‑thread reading for the cells collection (optional, improves read performance).
            wb.Worksheets[0].Cells.MultiThreadReading = true;

            // Import the XML data into the specified sheet starting at cell A1 (row 0, column 0).
            wb.ImportXml(job.XmlPath, job.SheetName, 0, 0);

            // Save the workbook to the designated file.
            wb.Save(job.OutputPath);
        });

        Console.WriteLine("All workbooks have been processed.");
    }
}