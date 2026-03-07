using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Drawing.Printing;

namespace AsposeCellsConsoleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Console Output Demo");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure image/print options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,          // each page on a separate sheet
                PrintWithStatusDialog = false    // suppress print status dialog
            };

            // Create SheetRender instance using the provided rule
            SheetRender render = new SheetRender(sheet, options);

            // Output rendering information to the console
            Console.WriteLine($"Total pages to print: {render.PageCount}");
            Console.WriteLine($"Page scale: {render.PageScale}");

            // Attempt to send the worksheet to the default printer
            try
            {
                // Use dynamic invocation to call the overload that accepts (PrinterSettings, string)
                // Passing null for PrinterSettings uses the system default printer
                dynamic dynRender = render;
                dynRender.ToPrinter(null, "ConsoleDemoJob");
                Console.WriteLine("Print job sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Printing failed: {ex.Message}");
            }
            finally
            {
                // Release resources
                render.Dispose();
            }
        }
    }
}