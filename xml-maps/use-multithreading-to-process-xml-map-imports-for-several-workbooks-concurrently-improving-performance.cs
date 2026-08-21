// Title: Parallel XML Map Import into Multiple Workbooks with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to use Parallel.ForEach to create separate Workbook instances, import XML data with ImportXml, and save each workbook concurrently. The example speeds up batch XML‑to‑Excel conversions by leveraging multithreading while keeping each workbook lifecycle isolated.
// Keywords: Aspose.Cells | C# | Parallel.ForEach | ImportXml | multithreaded Excel generation | XML map import | batch workbook processing | performance optimization | thread‑safe workbook creation
// Common Searches: Aspose.Cells parallel XML import example | C# import XML into multiple Excel files concurrently | How to use Parallel.ForEach with Aspose.Cells | Thread‑safe workbook creation Aspose.Cells | Improve XML map import speed .NET
// Developer Intent: Import XML data into several Excel workbooks at the same time using multithreading to reduce overall processing time.
// Use Cases: Convert a large set of XML feeds to individual Excel reports in parallel on a server. | Accelerate nightly batch jobs that generate workbooks from XML sources. | Scale real‑time data pipelines where each XML payload must be transformed into its own workbook.
// AI Prompts: Show a C# Aspose.Cells sample that uses Parallel.ForEach to import XML maps into multiple workbooks and includes per‑task error handling. | Explain how to control the degree of parallelism for large XML import batches to avoid memory pressure. | Provide code to log success or failure for each workbook created during a parallel XML import operation.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates how to use Parallel.ForEach to create separate Workbook instances, import XML data with ImportXml, and save each workbook concurrently. The example speeds up batch XML‑to‑Excel conversions by leveraging multithreading while keeping each workbook lifecycle isolated.
class XmlMapImportParallel
{
    static void Main()
    {
        // Define the XML source files, target sheet names and output workbook files.
        var jobs = new List<(string xmlPath, string outputPath, string sheetName)>
        {
            ("data1.xml", "Workbook1.xlsx", "Sheet1"),
            ("data2.xml", "Workbook2.xlsx", "Sheet1"),
            ("data3.xml", "Workbook3.xlsx", "Sheet1")
        };

        // Process each workbook concurrently.
        Parallel.ForEach(jobs, job =>
        {
            // Create a new workbook (lifecycle rule).
            Workbook wb = new Workbook();

            // Import the XML data into the specified sheet starting at cell A1 (rule).
            wb.ImportXml(job.xmlPath, job.sheetName, 0, 0);

            // Save the workbook to the designated file (rule).
            wb.Save(job.outputPath);
        });

        Console.WriteLine("All workbooks have been processed.");
    }
}
