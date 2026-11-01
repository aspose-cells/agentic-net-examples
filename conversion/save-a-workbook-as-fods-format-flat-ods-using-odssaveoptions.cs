// Title: Save an Aspose.Cells workbook as Flat ODS (FODS) using OdsSaveOptions in C#
// Common Searches: C# code example for saving an Excel workbook as Flat ODS with Aspose.Cells | How to use OdsSaveOptions to generate a .fods file in .NET | Aspose.Cells export to FODS format from a workbook programmatically | Convert Excel worksheet to Flat ODS using Aspose.Cells library in C# | Saving workbook as FODS using Aspose.Cells OdsSaveOptions tutorial
// Tags: Aspose.Cells OdsSaveOptions FODS export | C# generate Flat ODS file | Save workbook as .fods format | Flat ODS conversion with Aspose.Cells .NET | Aspose.Cells generate FODS output

// Create a new workbook
Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();

// (Optional) Add some data to the first worksheet
Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Sample data");

// Configure ODS save options for Flat ODS (FODS) format
Aspose.Cells.OdsSaveOptions odsOptions = new Aspose.Cells.OdsSaveOptions(Aspose.Cells.SaveFormat.FODS);

// Save the workbook as a FODS file
workbook.Save("output.fods", odsOptions);
