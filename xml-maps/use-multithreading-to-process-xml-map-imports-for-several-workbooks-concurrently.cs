using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Define a collection of XML import jobs.
        // Each job contains the XML file path, the output workbook file name, and the target sheet name.
        var jobs = new List<(string XmlPath, string OutputPath, string SheetName)>
        {
            ("data1.xml", "Workbook1.xlsx", "Sheet1"),
            ("data2.xml", "Workbook2.xlsx", "Sheet1"),
            ("data3.xml", "Workbook3.xlsx", "Sheet1")
        };

        // Process the jobs concurrently using Parallel.ForEach.
        Parallel.ForEach(jobs, job =>
        {
            // Create a new workbook instance.
            Workbook workbook = new Workbook();

            // Import the XML data into the specified sheet starting at cell A1 (row 0, column 0).
            workbook.ImportXml(job.XmlPath, job.SheetName, 0, 0);

            // Save the workbook to the designated output file.
            workbook.Save(job.OutputPath);
        });

        Console.WriteLine("All XML imports have been completed.");
    }
}