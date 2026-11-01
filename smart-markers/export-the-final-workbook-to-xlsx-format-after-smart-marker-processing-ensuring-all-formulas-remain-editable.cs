// Title: Save a workbook as XLSX after processing smart markers while keeping formulas editable using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an XLSX file (or creates a new workbook), assigns a DataTable as the data source, runs the smart‑marker engine, and saves the workbook as a new XLSX file while leaving all formulas editable. | Adapt the sample to accept a DataSet with several tables, apply smart markers to matching worksheets, and write each sheet to its own XLSX file without converting formulas to static values.
// Common Searches: how to retain formula editability after running Aspose.Cells smart markers in C# | C# export workbook to xlsx using Aspose.Cells designer without locking formulas | set calculation mode to manual after smart‑marker processing Aspose.Cells C# | initialize workbook for smart‑marker replacement when source file may be missing
// Tags: save processed workbook as xlsx Aspose.Cells | keep formulas editable after smart marker run C# | set calculation mode manual post‑designer Aspose.Cells | conditional workbook creation for smart markers C# | assign DataTable to WorkbookDesigner Aspose.Cells

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file or creates a new workbook, places a placeholder smart marker, assigns a DataTable as the data source for WorkbookDesigner, processes the smart markers, optionally sets manual calculation mode, ensures the output directory exists, and saves the final workbook as an XLSX file with all formulas remaining editable.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                // Placeholder cell for potential smart markers
                ws.Cells["A1"].PutValue("&=SmartMarker");
            }

            // Initialize designer for smart marker processing
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Sample data source
            DataTable data = new DataTable();
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Score", typeof(int));
            data.Rows.Add("Alice", 85);
            data.Rows.Add("Bob", 92);

            // Assign data source to the designer
            designer.SetDataSource(data);

            // Process smart markers
            designer.Process();

            // Optional: set manual calculation mode if supported by the library version
            // workbook.Settings.CalcMode = CalculationMode.Manual;

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the final workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
