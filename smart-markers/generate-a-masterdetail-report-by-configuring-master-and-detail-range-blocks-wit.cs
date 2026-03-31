using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace MasterDetailReport
{
    // Simple data models for master and detail
    public class Master
    {
        public string Category { get; set; }
        public List<Detail> Items { get; set; }
    }

    public class Detail
    {
        public string Product { get; set; }
        public int Quantity { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Paths for the template and the generated report
            string templatePath = "template.xlsx";   // XLSX file that contains the smart marker ranges
            string outputPath   = "report.xlsx";

            // Load the template workbook (create/load rule)
            Workbook workbook = new Workbook(templatePath);

            // -------------------------------------------------
            // Prepare sample master‑detail data
            // -------------------------------------------------
            var masters = new List<Master>
            {
                new Master
                {
                    Category = "Electronics",
                    Items = new List<Detail>
                    {
                        new Detail { Product = "Laptop",   Quantity = 5 },
                        new Detail { Product = "Smartphone", Quantity = 12 }
                    }
                },
                new Master
                {
                    Category = "Furniture",
                    Items = new List<Detail>
                    {
                        new Detail { Product = "Desk",   Quantity = 3 },
                        new Detail { Product = "Chair",  Quantity = 8 },
                        new Detail { Product = "Cabinet",Quantity = 2 }
                    }
                }
            };

            // -------------------------------------------------
            // Configure the WorkbookDesigner (smart markers)
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Set the master data source. In the template the master range should be named "_Master"
            designer.SetDataSource("Master", masters);

            // Set the detail data source. In the template the detail range should be named "_Detail"
            // Aspose.Cells automatically links the detail rows to the current master row.
            designer.SetDataSource("Detail", masters[0].Items); // placeholder, will be overridden per master

            // Process the smart markers. The designer will repeat the master range and, for each master row,
            // repeat the detail range using the corresponding Items collection.
            designer.Process();

            // Save the generated report (save rule)
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}