using System;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

class ChartSheetPdfBookmarksDemo
{
    static void Main()
    {
        SimpleBookmarkDemo();
        HierarchicalBookmarksDemo();
        CollapsibleBookmarksDemo();
        NamedDestinationBookmarksDemo();
        SheetSetWithBookmarksDemo();
    }

    // 1. Simple bookmark that points to a cell near a chart on a single sheet
    static void SimpleBookmarkDemo()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Sales";

        ws.Cells["A1"].PutValue("Month");
        ws.Cells["B1"].PutValue("Revenue");
        ws.Cells["A2"].PutValue("Jan");
        ws.Cells["B2"].PutValue(1200);
        ws.Cells["A3"].PutValue("Feb");
        ws.Cells["B3"].PutValue(1500);
        ws.Cells["A4"].PutValue("Mar");
        ws.Cells["B4"].PutValue(1800);

        int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Quarterly Revenue";

        PdfBookmarkEntry bookmark = new PdfBookmarkEntry
        {
            Text = "Sales Chart",
            Destination = ws.Cells["A1"],
            IsOpen = true
        };

        PdfSaveOptions options = new PdfSaveOptions { Bookmark = bookmark };
        wb.Save("SimpleBookmark.pdf", options);
    }

    // 2. Hierarchical bookmarks: root entry with child entries for each sheet's chart
    static void HierarchicalBookmarksDemo()
    {
        Workbook wb = new Workbook();

        // Sheet 1 – Products
        Worksheet ws1 = wb.Worksheets[0];
        ws1.Name = "Products";
        ws1.Cells["A1"].PutValue("Product");
        ws1.Cells["B1"].PutValue("Units");
        ws1.Cells["A2"].PutValue("A");
        ws1.Cells["B2"].PutValue(500);
        ws1.Cells["A3"].PutValue("B");
        ws1.Cells["B3"].PutValue(300);
        int c1 = ws1.Charts.Add(ChartType.Pie, 5, 0, 15, 10);
        Chart chart1 = ws1.Charts[c1];
        chart1.NSeries.Add("B2:B3", true);
        chart1.NSeries.CategoryData = "A2:A3";
        chart1.Title.Text = "Product Distribution";

        // Sheet 2 – Expenses
        Worksheet ws2 = wb.Worksheets.Add("Expenses");
        ws2.Cells["A1"].PutValue("Category");
        ws2.Cells["B1"].PutValue("Amount");
        ws2.Cells["A2"].PutValue("Rent");
        ws2.Cells["B2"].PutValue(2000);
        ws2.Cells["A3"].PutValue("Utilities");
        ws2.Cells["B3"].PutValue(800);
        int c2 = ws2.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart2 = ws2.Charts[c2];
        chart2.NSeries.Add("B2:B3", true);
        chart2.NSeries.CategoryData = "A2:A3";
        chart2.Title.Text = "Monthly Expenses";

        // Sheet 3 – Profit
        Worksheet ws3 = wb.Worksheets.Add("Profit");
        ws3.Cells["A1"].PutValue("Month");
        ws3.Cells["B1"].PutValue("Profit");
        ws3.Cells["A2"].PutValue("Jan");
        ws3.Cells["B2"].PutValue(300);
        ws3.Cells["A3"].PutValue("Feb");
        ws3.Cells["B3"].PutValue(700);
        int c3 = ws3.Charts.Add(ChartType.Line, 5, 0, 15, 10);
        Chart chart3 = ws3.Charts[c3];
        chart3.NSeries.Add("B2:B3", true);
        chart3.NSeries.CategoryData = "A2:A3";
        chart3.Title.Text = "Profit Trend";

        // Build bookmark hierarchy
        PdfBookmarkEntry root = new PdfBookmarkEntry
        {
            Text = "Report Overview",
            Destination = ws1.Cells["A1"],
            IsOpen = true,
            SubEntry = new ArrayList()
        };
        root.SubEntry.Add(new PdfBookmarkEntry { Text = "Products Chart", Destination = ws1.Cells["A1"] });
        root.SubEntry.Add(new PdfBookmarkEntry { Text = "Expenses Chart", Destination = ws2.Cells["A1"] });
        root.SubEntry.Add(new PdfBookmarkEntry { Text = "Profit Chart", Destination = ws3.Cells["A1"] });

        PdfSaveOptions options = new PdfSaveOptions { Bookmark = root };
        wb.Save("HierarchicalBookmarks.pdf", options);
    }

    // 3. Collapsible sections using IsCollapse / IsOpen
    static void CollapsibleBookmarksDemo()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Analytics";

        // First half data
        ws.Cells["A1"].PutValue("Quarter");
        ws.Cells["B1"].PutValue("Sales");
        ws.Cells["A2"].PutValue("Q1");
        ws.Cells["B2"].PutValue(4000);
        ws.Cells["A3"].PutValue("Q2");
        ws.Cells["B3"].PutValue(4500);
        int chartIdx1 = ws.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        Chart chart1 = ws.Charts[chartIdx1];
        chart1.NSeries.Add("B2:B3", true);
        chart1.NSeries.CategoryData = "A2:A3";
        chart1.Title.Text = "First Half Sales";

        // Second half data
        ws.Cells["A5"].PutValue("Quarter");
        ws.Cells["B5"].PutValue("Sales");
        ws.Cells["A6"].PutValue("Q3");
        ws.Cells["B6"].PutValue(4700);
        ws.Cells["A7"].PutValue("Q4");
        ws.Cells["B7"].PutValue(5200);
        int chartIdx2 = ws.Charts.Add(ChartType.Column, 20, 0, 30, 10);
        Chart chart2 = ws.Charts[chartIdx2];
        chart2.NSeries.Add("B6:B7", true);
        chart2.NSeries.CategoryData = "A6:A7";
        chart2.Title.Text = "Second Half Sales";

        // Hidden root (children appear at top level)
        PdfBookmarkEntry root = new PdfBookmarkEntry
        {
            Text = null,
            SubEntry = new ArrayList()
        };

        // First half bookmark (collapsed)
        PdfBookmarkEntry firstHalf = new PdfBookmarkEntry
        {
            Text = "First Half",
            Destination = ws.Cells["A1"],
            IsCollapse = true
        };

        // Second half bookmark (expanded)
        PdfBookmarkEntry secondHalf = new PdfBookmarkEntry
        {
            Text = "Second Half",
            Destination = ws.Cells["A5"],
            IsOpen = true
        };

        root.SubEntry.Add(firstHalf);
        root.SubEntry.Add(secondHalf);

        PdfSaveOptions options = new PdfSaveOptions { Bookmark = root };
        wb.Save("CollapsibleBookmarks.pdf", options);
    }

    // 4. Using DestinationName to create a named destination
    static void NamedDestinationBookmarksDemo()
    {
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Name = "Summary";

        ws.Cells["A1"].PutValue("Metric");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("Total");
        ws.Cells["B2"].PutValue(12345);
        int chartIdx = ws.Charts.Add(ChartType.Pie, 5, 0, 15, 10);
        Chart chart = ws.Charts[chartIdx];
        chart.NSeries.Add("B2:B2", true);
        chart.Title.Text = "Total Metric";

        PdfBookmarkEntry entry = new PdfBookmarkEntry
        {
            Text = "Summary Section",
            Destination = ws.Cells["A1"],
            DestinationName = "SummaryStart",
            IsOpen = true
        };

        PdfSaveOptions options = new PdfSaveOptions { Bookmark = entry };
        wb.Save("NamedDestinationBookmark.pdf", options);
    }

    // 5. Exporting a subset of sheets while preserving bookmarks
    static void SheetSetWithBookmarksDemo()
    {
        Workbook wb = new Workbook();

        // Sheet A
        Worksheet wsA = wb.Worksheets[0];
        wsA.Name = "SheetA";
        wsA.Cells["A1"].PutValue("Data A1");
        int chartA = wsA.Charts.Add(ChartType.Column, 5, 0, 15, 10);
        wsA.Charts[chartA].NSeries.Add("A1:A1", true);
        wsA.Charts[chartA].Title.Text = "Chart A";

        // Sheet B
        Worksheet wsB = wb.Worksheets.Add("SheetB");
        wsB.Cells["A1"].PutValue("Data B1");
        int chartB = wsB.Charts.Add(ChartType.Pie, 5, 0, 15, 10);
        wsB.Charts[chartB].NSeries.Add("A1:A1", true);
        wsB.Charts[chartB].Title.Text = "Chart B";

        // Sheet C (will be excluded from PDF)
        Worksheet wsC = wb.Worksheets.Add("SheetC");
        wsC.Cells["A1"].PutValue("Data C1");
        int chartC = wsC.Charts.Add(ChartType.Line, 5, 0, 15, 10);
        wsC.Charts[chartC].NSeries.Add("A1:A1", true);
        wsC.Charts[chartC].Title.Text = "Chart C";

        // Bookmarks for the exported sheets only
        PdfBookmarkEntry root = new PdfBookmarkEntry
        {
            Text = "Workbook Sections",
            Destination = wsA.Cells["A1"],
            IsOpen = true,
            SubEntry = new ArrayList()
        };
        root.SubEntry.Add(new PdfBookmarkEntry { Text = "Sheet A", Destination = wsA.Cells["A1"] });
        root.SubEntry.Add(new PdfBookmarkEntry { Text = "Sheet B", Destination = wsB.Cells["A1"] });

        PdfSaveOptions options = new PdfSaveOptions
        {
            Bookmark = root,
            SheetSet = new SheetSet(new int[] { 0, 1 }) // export only SheetA and SheetB
        };

        wb.Save("SheetSetBookmarks.pdf", options);
    }
}