// Title: Identify and Fix Corrupted Named Ranges Caused by Worksheet Renaming with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to scan all workbook defined names and replace any RefersTo reference that points to a missing worksheet with a reference to the first worksheet, preserving the original cell address. | Generate a method that validates named ranges in an Excel file and automatically corrects those whose sheet name no longer exists after a rename, using the Aspose.Cells .NET API. | Create a script that loads an .xlsx file, detects named ranges referencing deleted sheets, updates their RefersTo strings to a fallback sheet, and saves the workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# fix named range after sheet name change | how to update RefersTo for corrupted named ranges in .NET | detect missing worksheet references in Excel named ranges using Aspose | C# code to reassign named ranges to a default sheet when original sheet is removed | automate correction of named ranges that point to deleted sheets with Aspose.Cells
// Tags: named range corruption detection Aspose.Cells | refersTo sheet reference correction C# | validate defined names after worksheet rename | fallback sheet assignment for missing named ranges | Aspose.Cells workbook named range repair

// Load the workbook (using the provided load rule)
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook("input.xlsx");

// Build a set of existing worksheet names for quick lookup
System.Collections.Generic.HashSet<string> sheetNames = new System.Collections.Generic.HashSet<string>();
foreach (Aspose.Cells.Worksheet sheet in workbook.Worksheets)
{
    sheetNames.Add(sheet.Name);
}

// Reference to the first worksheet (used for correction if needed)
Aspose.Cells.Worksheet firstSheet = workbook.Worksheets[0];
string firstSheetName = firstSheet.Name;

// Iterate through all defined names (named ranges)
foreach (Aspose.Cells.Name definedName in workbook.Worksheets.Names)
{
    // The RefersTo string is like "=Sheet1!$A$1:$B$2"
    string refersTo = definedName.RefersTo;

    // Ensure the string contains a sheet reference
    if (string.IsNullOrEmpty(refersTo) || !refersTo.StartsWith("=") || !refersTo.Contains("!"))
        continue; // Not a standard range reference, skip

    // Extract the sheet name part (between '=' and '!')
    int exclPos = refersTo.IndexOf('!');
    string sheetNameInRef = refersTo.Substring(1, exclPos - 1); // exclude leading '='

    // Check if the referenced sheet still exists
    if (!sheetNames.Contains(sheetNameInRef))
    {
        // The named range is corrupted because its sheet no longer exists
        // Correct it by pointing to the same cell address on the first worksheet
        string addressPart = refersTo.Substring(exclPos); // includes '!' and the range
        string newRefersTo = "=" + firstSheetName + addressPart;

        // Apply the correction
        definedName.RefersTo = newRefersTo;
    }
}

// Save the corrected workbook (using the provided save rule)
workbook.Save("output.xlsx");
