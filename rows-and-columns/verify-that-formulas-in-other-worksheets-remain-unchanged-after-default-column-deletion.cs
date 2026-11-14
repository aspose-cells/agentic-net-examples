// Title: Verify unchanged formulas in other worksheets after deleting a column with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if formulas are affected when removing column A from a worksheet | compare worksheet formulas before and after column deletion using Aspose.Cells | detect formula changes in Excel sheets after deleting the default column with Aspose.Cells | how to ensure formulas in other worksheets stay the same after column removal in Aspose.Cells | validate Excel formulas across multiple sheets after column delete in .NET
// Tags: Aspose.Cells column deletion formula integrity | C# Excel formula consistency after column removal | Aspose.Cells track formula stability across sheets | validate worksheet formula unchanged state .NET | Excel column delete impact analysis Aspose.Cells

// Load an existing workbook (replace with your actual file path)
var workbook = new Aspose.Cells.Workbook("input.xlsx");

// Store formulas from all worksheets before any modification
var originalFormulas = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>();
foreach (Aspose.Cells.Worksheet ws in workbook.Worksheets)
{
    var formulas = new System.Collections.Generic.List<string>();
    foreach (Aspose.Cells.Cell cell in ws.Cells)
    {
        if (!string.IsNullOrEmpty(cell.Formula))
        {
            // Store cell address and its formula for exact comparison later
            formulas.Add(cell.Name + ":" + cell.Formula);
        }
    }
    originalFormulas[ws.Name] = formulas;
}

// Delete the default column (e.g., column A) in the first worksheet
Aspose.Cells.Worksheet sheetToModify = workbook.Worksheets[0];
sheetToModify.Cells.DeleteColumn(0); // 0‑based index, deletes column A

// Store formulas from all worksheets after the column deletion
var afterFormulas = new System.Collections.Generic.Dictionary<string, System.Collections.Generic.List<string>>();
foreach (Aspose.Cells.Worksheet ws in workbook.Worksheets)
{
    var formulas = new System.Collections.Generic.List<string>();
    foreach (Aspose.Cells.Cell cell in ws.Cells)
    {
        if (!string.IsNullOrEmpty(cell.Formula))
        {
            formulas.Add(cell.Name + ":" + cell.Formula);
        }
    }
    afterFormulas[ws.Name] = formulas;
}

// Verify that formulas in worksheets other than the modified one remain unchanged
bool formulasUnchanged = true;
foreach (var kvp in originalFormulas)
{
    string sheetName = kvp.Key;
    // Skip the worksheet where the column was deleted
    if (sheetName == sheetToModify.Name) continue;

    var beforeList = kvp.Value;
    var afterList = afterFormulas[sheetName];

    // Quick count check
    if (beforeList.Count != afterList.Count)
    {
        formulasUnchanged = false;
        break;
    }

    // Detailed element‑by‑element comparison
    for (int i = 0; i < beforeList.Count; i++)
    {
        if (beforeList[i] != afterList[i])
        {
            formulasUnchanged = false;
            break;
        }
    }

    if (!formulasUnchanged) break;
}

// Output verification result
System.Console.WriteLine(formulasUnchanged
    ? "All formulas in other worksheets remain unchanged after column deletion."
    : "Some formulas in other worksheets have changed after column deletion.");

// Save the workbook if further inspection is needed
workbook.Save("output.xlsx");
