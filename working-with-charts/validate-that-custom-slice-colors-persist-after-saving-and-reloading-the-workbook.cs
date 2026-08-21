// Title: Verify Slicer Style Persistence After Saving Workbook with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a pivot table, insert a slicer, apply the built‑in SlicerStyleDark2, save the file, reload it, and confirm that the slicer's StyleType remains unchanged, proving that custom slicer colors survive serialization in Aspose.Cells.
// Keywords: Aspose.Cells | C# | slicer style persistence | SlicerStyleDark2 | Excel slicer validation | pivot table slicer | save and reload workbook | style property verification | Excel automation testing
// Common Searches: Aspose.Cells verify slicer style after save | C# check slicer color persistence in Excel | how to test slicer style retention with Aspose.Cells | validate pivot slicer appearance after workbook reload | ensure custom slicer colors survive serialization .NET
// Developer Intent: Confirm that a slicer's custom style (color) is retained after the workbook is saved and reopened using Aspose.Cells for .NET.
// Use Cases: Automated regression tests for Excel reports that rely on specific slicer colors. | CI/CD pipelines that validate UI consistency of generated workbooks. | Quality assurance checks before distributing programmatically created Excel files to end users.
// AI Prompts: Generate a C# function that creates a slicer, sets SlicerStyleDark2, saves the workbook, reloads it, and asserts the style persisted with Aspose.Cells. | Write an NUnit test case to verify slicer style persistence after workbook serialization using Aspose.Cells for .NET. | Explain the internal storage of slicer style information in an XLSX file and why Aspose.Cells preserves it across save/load operations.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Slicers; // for SlicerStyleType

// Create a new workbook
Workbook workbook = new Workbook();

// ---------- Setup data ----------
Worksheet dataSheet = workbook.Worksheets[0];
dataSheet.Name = "Data";
dataSheet.Cells["A1"].PutValue("Category");
dataSheet.Cells["A2"].PutValue("Apple");
dataSheet.Cells["A3"].PutValue("Banana");
dataSheet.Cells["A4"].PutValue("Orange");
dataSheet.Cells["B1"].PutValue("Sales");
dataSheet.Cells["B2"].PutValue(120);
dataSheet.Cells["B3"].PutValue(150);
dataSheet.Cells["B4"].PutValue(90);

// ---------- Create a pivot table ----------
Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

// ---------- Add a slicer ----------
Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
int slicerIdx = slicerSheet.Slicers.Add(pivot, "A1", "Category");
Slicer slicer = slicerSheet.Slicers[slicerIdx];

// Set a built‑in slicer style (this determines the slicer colors)
slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

// ---------- Save the workbook ----------
string filePath = "SlicerCustomColorDemo.xlsx";
workbook.Save(filePath);

// ---------- Reload and validate ----------
Workbook loadedWb = new Workbook(filePath);
Worksheet loadedSlicerSheet = loadedWb.Worksheets["Slicer"];
Slicer loadedSlicer = loadedSlicerSheet.Slicers[0];

// Verify that the slicer style persisted after reload
bool stylePersisted = loadedSlicer.StyleType == SlicerStyleType.SlicerStyleDark2;
Console.WriteLine("Slicer style persisted: " + stylePersisted);

// Optional: output the style type for confirmation
Console.WriteLine("Loaded slicer style: " + loadedSlicer.StyleType);
