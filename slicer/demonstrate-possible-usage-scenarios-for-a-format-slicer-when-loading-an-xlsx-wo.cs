using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class FormatSlicerDemo
{
    static void Main()
    {
        // Paths to the source XLSX file and output files
        string inputPath = "input.xlsx";
        string outputPathAll = "output_all.xlsx";
        string outputPathSlicersOnly = "output_slicers_only.xlsx";

        // -------------------------------------------------
        // Scenario 1: Load the workbook with default settings (auto format detection)
        // -------------------------------------------------
        Workbook wbDefault = new Workbook(inputPath);
        ShowSlicerInfo(wbDefault);
        wbDefault.Save(outputPathAll);

        // -------------------------------------------------
        // Scenario 2: Load the workbook explicitly specifying the Xlsx format
        // -------------------------------------------------
        LoadOptions optExplicit = new LoadOptions(LoadFormat.Xlsx);
        Workbook wbExplicit = new Workbook(inputPath, optExplicit);
        ShowSlicerInfo(wbExplicit);
        wbExplicit.Save(outputPathAll); // reuse the same output file for demonstration

        // -------------------------------------------------
        // Scenario 3: Load only slicer‑related objects using a custom LoadFilter
        // -------------------------------------------------
        LoadOptions optSlicerOnly = new LoadOptions(LoadFormat.Xlsx);
        optSlicerOnly.LoadFilter = new SlicerOnlyLoadFilter(); // custom filter defined below
        Workbook wbSlicerOnly = new Workbook(inputPath, optSlicerOnly);
        ShowSlicerInfo(wbSlicerOnly);
        wbSlicerOnly.Save(outputPathSlicersOnly);
    }

    // Helper method to enumerate slicers in each worksheet and display basic properties
    static void ShowSlicerInfo(Workbook wb)
    {
        foreach (Worksheet ws in wb.Worksheets)
        {
            Console.WriteLine($"Worksheet: {ws.Name}, Slicer count: {ws.Slicers.Count}");
            for (int i = 0; i < ws.Slicers.Count; i++)
            {
                Slicer slicer = ws.Slicers[i];
                Console.WriteLine($"  Slicer {i}: Name=\"{slicer.Name}\", Caption=\"{slicer.Caption}\", Style={slicer.StyleType}");
            }
        }
    }

    // Custom LoadFilter that loads only drawing objects (which include slicers) and sheet settings
    class SlicerOnlyLoadFilter : LoadFilter
    {
        public override void StartSheet(Worksheet sheet)
        {
            // Load only the drawing objects (slicers are drawing objects) and basic sheet settings
            LoadDataFilterOptions = LoadDataFilterOptions.Drawing | LoadDataFilterOptions.SheetSettings;
        }
    }
}