using System;
using System.Diagnostics;
using Aspose.Cells;

namespace PageSetupCloneBenchmark
{
    class Program
    {
        static void Main()
        {
            // ---------- Prepare a source PageSetup with several properties ----------
            Workbook sourceWb = new Workbook();                         // create workbook
            Worksheet sourceWs = sourceWb.Worksheets[0];
            PageSetup sourcePs = sourceWs.PageSetup;
            sourcePs.PaperSize = PaperSizeType.PaperA4;
            sourcePs.Orientation = PageOrientationType.Portrait;
            sourcePs.FitToPagesWide = 1;
            sourcePs.FitToPagesTall = 0;
            sourcePs.PrintGridlines = true;
            sourcePs.Zoom = 100;

            // ---------- Benchmark cloning via PageSetup.Copy ----------
            GC.Collect();
            long memBeforeClone = GC.GetTotalMemory(true);

            for (int i = 1; i <= 50; i++)
            {
                // Add a new worksheet
                Worksheet ws = sourceWb.Worksheets.Add($"Clone{i}");
                // Clone the page setup from the source worksheet
                ws.PageSetup.Copy(sourcePs, new CopyOptions());
            }

            GC.Collect();
            long memAfterClone = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory used after cloning (Copy): {memAfterClone - memBeforeClone:N0} bytes");

            // Save the workbook that used cloning (optional)
            sourceWb.Save("ClonedPageSetup.xlsx");

            // ---------- Benchmark direct property assignment ----------
            Workbook assignWb = new Workbook();                         // create a second workbook
            Worksheet assignSourceWs = assignWb.Worksheets[0];
            PageSetup assignSourcePs = assignSourceWs.PageSetup;
            // Apply the same initial settings
            assignSourcePs.PaperSize = sourcePs.PaperSize;
            assignSourcePs.Orientation = sourcePs.Orientation;
            assignSourcePs.FitToPagesWide = sourcePs.FitToPagesWide;
            assignSourcePs.FitToPagesTall = sourcePs.FitToPagesTall;
            assignSourcePs.PrintGridlines = sourcePs.PrintGridlines;
            assignSourcePs.Zoom = sourcePs.Zoom;

            GC.Collect();
            long memBeforeAssign = GC.GetTotalMemory(true);

            for (int i = 1; i <= 50; i++)
            {
                Worksheet ws = assignWb.Worksheets.Add($"Assign{i}");
                PageSetup ps = ws.PageSetup;
                // Manually copy each property
                ps.PaperSize = assignSourcePs.PaperSize;
                ps.Orientation = assignSourcePs.Orientation;
                ps.FitToPagesWide = assignSourcePs.FitToPagesWide;
                ps.FitToPagesTall = assignSourcePs.FitToPagesTall;
                ps.PrintGridlines = assignSourcePs.PrintGridlines;
                ps.Zoom = assignSourcePs.Zoom;
            }

            GC.Collect();
            long memAfterAssign = GC.GetTotalMemory(true);
            Console.WriteLine($"Memory used after direct assignment: {memAfterAssign - memBeforeAssign:N0} bytes");

            // Save the workbook that used direct assignment (optional)
            assignWb.Save("AssignedPageSetup.xlsx");
        }
    }
}