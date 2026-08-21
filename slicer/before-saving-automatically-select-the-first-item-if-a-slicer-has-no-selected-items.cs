// Title: Auto‑select first slicer item when none are selected – Aspose.Cells for .NET (C#) example
// Description: C# snippet that loads a workbook, scans every worksheet for slicers, checks each slicer's cache items and automatically marks the first item as selected if the slicer has no active selection, then saves the workbook. Ideal for ensuring default filter values in Excel reports generated with Aspose.Cells.
// Keywords: Aspose.Cells slicer selection | C# default slicer item | programmatic slicer selection .NET | Excel slicer cache items | auto select slicer first item | Aspose.Cells workbook save | pivot table slicer default | Excel automation Aspose.Cells | GitHub Aspose.Cells example | coding‑agent slicer utility
// Common Searches: how to set default slicer selection using Aspose.Cells | Aspose.Cells C# select first slicer item if none selected | auto select slicer item before saving workbook | Aspose.Cells iterate slicers and set selection | C# code to ensure slicer has a selected value
// Developer Intent: Programmatically guarantee that every slicer in an Excel workbook has at least one selected value by automatically selecting the first cache item when the slicer is empty, then persist the changes.
// Use Cases: Generate recurring reports where a slicer must always have a fallback value to prevent empty result sets. | Prepare workbooks for PDF or image export, ensuring pivot tables reflect a valid slicer filter. | Build data‑driven dashboards that automatically apply a default filter when users clear all slicer selections.
// AI Prompts: Create a reusable C# method that scans all slicers in an Aspose.Cells workbook and selects the first cache item when no items are selected. | Provide an Aspose.Cells for .NET example that checks slicer selections, sets items[0].Selected = true if needed, and saves the workbook. | Generate a utility class for Aspose.Cells that enforces a default slicer selection across multiple worksheets and can be called before exporting.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerSelectionDemo
{
    // C# snippet that loads a workbook, scans every worksheet for slicers, checks each slicer's cache items and automatically marks the first item as selected if the slicer has no active selection, then saves the workbook. Ideal for ensuring default filter values in Excel reports generated with Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all slicers on the worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Get the collection of slicer cache items
                    SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

                    // Determine if any item is already selected
                    bool anySelected = false;
                    for (int i = 0; i < items.Count; i++)
                    {
                        if (items[i].Selected)
                        {
                            anySelected = true;
                            break;
                        }
                    }

                    // If no items are selected, select the first item (if any exist)
                    if (!anySelected && items.Count > 0)
                    {
                        items[0].Selected = true;
                    }
                }
            }

            // Save the workbook after processing slicers
            workbook.Save("output.xlsx");
        }
    }
}
