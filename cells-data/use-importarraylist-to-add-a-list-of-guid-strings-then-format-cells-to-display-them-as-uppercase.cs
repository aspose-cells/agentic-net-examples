// Title: Import an ArrayList of uppercase GUID strings into an Excel worksheet and apply Calibri styling with Aspose.Cells for .NET
// AI Prompts: Generate five GUID values, convert each to an uppercase string, add them to a System.Collections.ArrayList, and import the list into cell A1 of the first worksheet using ImportArrayList. | Create a Calibri font style, apply it to the range that contains the imported GUIDs, and save the workbook as GuidList.xlsx.
// Common Searches: Aspose.Cells C# ImportArrayList uppercase GUID list into Excel | How to format imported GUID cells as uppercase with Aspose.Cells | Applying custom font style to a range after ImportArrayList in Aspose.Cells | Saving a workbook that contains GUID strings using Aspose.Cells for .NET | ImportArrayList from System.Collections.ArrayList into worksheet Aspose.Cells example
// Tags: ImportArrayList GUID list Aspose.Cells | uppercase GUID cell formatting Aspose.Cells | apply Calibri style to imported range | save workbook as Excel file Aspose.Cells | generate GUID strings C# Aspose.Cells

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// Get the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];

// Prepare an ArrayList of GUID strings
System.Collections.ArrayList guidList = new System.Collections.ArrayList();
for (int i = 0; i < 5; i++)
{
    // Generate a new GUID and store it as an uppercase string
    guidList.Add(System.Guid.NewGuid().ToString().ToUpper());
}

// Import the ArrayList into the worksheet starting at cell A1 (row 0, column 0)
sheet.Cells.ImportArrayList(guidList, 0, 0, false);

// Apply a style to the imported range (optional visual formatting)
Aspose.Cells.Style style = workbook.CreateStyle();
style.Font.Name = "Calibri";
style.Font.Size = 11;

// Apply the style to the range that contains the GUIDs
int rowCount = guidList.Count;
Aspose.Cells.Range range = sheet.Cells.CreateRange(0, 0, rowCount, 1);
range.ApplyStyle(style, new Aspose.Cells.StyleFlag() { All = true });

// Save the workbook to a file
workbook.Save("GuidList.xlsx");
