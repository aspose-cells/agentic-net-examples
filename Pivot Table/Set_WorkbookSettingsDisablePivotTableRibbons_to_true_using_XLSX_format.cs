using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisablePivotTableRibbonsDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // (Optional) If the Aspose.Cells version supports it, you can disable PivotTable ribbons here.
            // workbook.Settings.EnablePivotTableRibbons = false; // Removed due to unavailable API in current version.

            // Save the workbook
            workbook.Save("DisablePivotTableRibbonsDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            DisablePivotTableRibbonsDemo.Run();
        }
    }
}