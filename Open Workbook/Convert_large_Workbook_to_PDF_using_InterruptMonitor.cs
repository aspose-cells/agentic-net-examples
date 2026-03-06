using System;
using System.IO;
using Aspose.Cells;

class ConvertLargeWorkbookToPdfWithInterrupt
{
    static void Main()
    {
        string inputPath = "LargeWorkbook.xlsx";
        string outputPath = "LargeWorkbook.pdf";

        if (!File.Exists(inputPath))
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            for (int i = 0; i < 5000; i++)
            {
                ws.Cells[i, 0].PutValue($"Row {i + 1}");
                ws.Cells[i, 1].PutValue(i);
            }
            wb.Save(inputPath);
        }

        ThreadInterruptMonitor monitor = new ThreadInterruptMonitor(false);

        LoadOptions loadOptions = new LoadOptions
        {
            InterruptMonitor = monitor
        };

        monitor.StartMonitor(5000);
        Workbook workbook = null;
        try
        {
            workbook = new Workbook(inputPath, loadOptions);
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Loading was interrupted.");
            return;
        }
        finally
        {
            monitor.FinishMonitor();
        }

        workbook.InterruptMonitor = monitor;

        monitor.StartMonitor(8000);
        try
        {
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine("Workbook saved to PDF successfully.");
        }
        catch (CellsException ex) when (ex.Code == ExceptionType.Interrupted)
        {
            Console.WriteLine("Saving was interrupted due to time limit.");
        }
        finally
        {
            monitor.FinishMonitor();
        }
    }
}