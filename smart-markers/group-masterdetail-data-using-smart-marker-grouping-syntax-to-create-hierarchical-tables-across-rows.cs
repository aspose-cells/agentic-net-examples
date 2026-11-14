// Title: Generate hierarchical master‑detail Excel tables with Aspose.Cells smart marker grouping in C#
// Common Searches: Aspose.Cells C# smart markers group master detail rows | How to use TableStart and GroupStart markers for hierarchical data in Aspose.Cells | C# example for master‑detail report using WorkbookDesigner | Create grouped master‑detail Excel sheet with smart marker syntax in .NET | Aspose.Cells smart marker grouping syntax for nested collections C# | Generate Excel master‑detail tables from List<Master> using Aspose.Cells
// Tags: Aspose.Cells smart marker grouping | C# master detail Excel report | WorkbookDesigner SetDataSource hierarchical data | TableStart TableEnd smart marker syntax | GroupStart GroupEnd Aspose.Cells | Excel hierarchical tables C#

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample creates a new workbook, inserts smart marker tags (TableStart, GroupStart, Master.Name, GroupEnd, Master.Details.Description, TableEnd) to define a master‑detail hierarchy, binds a List<Master> (each with a List<Detail>) to the WorkbookDesigner via SetDataSource, processes the markers, and saves the hierarchical Excel file as GroupedMasterDetail.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];

        // Add headers
        sheet.Cells["A1"].PutValue("Master Name");
        sheet.Cells["B1"].PutValue("Detail Description");

        // Insert smart markers using grouping syntax
        // Table start
        sheet.Cells["A2"].PutValue("&=TableStart");
        // Group start for master records
        sheet.Cells["A3"].PutValue("&=GroupStart");
        // Master field
        sheet.Cells["A4"].PutValue("&=Master.Name");
        // Group end
        sheet.Cells["A5"].PutValue("&=GroupEnd");
        // Detail field (will repeat for each detail of the current master)
        sheet.Cells["B6"].PutValue("&=Master.Details.Description");
        // Table end
        sheet.Cells["A7"].PutValue("&=TableEnd");

        // Prepare hierarchical data source
        var masters = new List<Master>();

        var master1 = new Master
        {
            Name = "Master 1",
            Details = new List<Detail>
            {
                new Detail { Description = "Detail 1‑A" },
                new Detail { Description = "Detail 1‑B" }
            }
        };
        masters.Add(master1);

        var master2 = new Master
        {
            Name = "Master 2",
            Details = new List<Detail>
            {
                new Detail { Description = "Detail 2‑A" },
                new Detail { Description = "Detail 2‑B" },
                new Detail { Description = "Detail 2‑C" }
            }
        };
        masters.Add(master2);

        // Process smart markers with WorkbookDesigner
        var designer = new WorkbookDesigner(workbook);
        designer.SetDataSource("Master", masters);
        designer.Process();

        // Save the result
        workbook.Save("GroupedMasterDetail.xlsx");
    }
}

// Master class containing a collection of details
public class Master
{
    public string Name { get; set; }
    public List<Detail> Details { get; set; }
}

// Detail class
public class Detail
{
    public string Description { get; set; }
}
