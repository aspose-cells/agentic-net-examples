using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

class XmlMapConcurrentImporter
{
    static void Main()
    {
        // Define the XML files to import and the corresponding output Excel files.
        var jobs = new List<(string XmlPath, string OutputPath)>
        {
            ("data1.xml", "output1.xlsx"),
            ("data2.xml", "output2.xlsx"),
            ("data3.xml", "output3.xlsx")
        };

        // Process each workbook in parallel to improve performance.
        Parallel.ForEach(jobs, job =>
        {
            // Create a new workbook instance.
            Workbook wb = new Workbook();

            // Import the XML data into the first worksheet starting at cell A1.
            wb.ImportXml(job.XmlPath, "Sheet1", 0, 0);

            // Save the workbook to the designated file.
            wb.Save(job.OutputPath);
        });

        Console.WriteLine("All workbooks have been processed.");
    }
}