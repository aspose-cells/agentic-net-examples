// Title: Show rows only when a numeric field exceeds a threshold using an IF smart marker with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel template containing a smart marker like ${if:Amount>100}${Amount}${endif}, binds a DataTable as the data source, processes the markers with WorkbookDesigner, and saves the filtered workbook. | Adapt an existing Aspose.Cells workbook to include a conditional smart marker that outputs a column only when its value is greater than a specified limit, using WorkbookDesigner in C#.
// Common Searches: Aspose.Cells WorkbookDesigner IF smart marker example C# | filter Excel rows with smart markers based on numeric value in .NET | how to use ${if:field>value} syntax in Aspose.Cells templates | conditional data rendering in Excel using Aspose.Cells smart markers | display only rows where Amount > 100 with Aspose.Cells smart markers
// Tags: Aspose.Cells IF smart marker | WorkbookDesigner conditional row output | C# smart marker numeric threshold | Excel template conditional data rendering | Aspose.Cells data source filtering

// Create a DataTable with sample data
System.Data.DataTable dt = new System.Data.DataTable();
dt.Columns.Add("Amount", typeof(double));
dt.Rows.Add(50);   // Below threshold, will not be displayed
dt.Rows.Add(150);  // Above threshold, will be displayed
dt.Rows.Add(250);  // Above threshold, will be displayed

// Load the Excel template that contains the smart marker
// Template should have a smart marker like: ${if:Amount>100}${Amount}${endif}
Aspose.Cells.WorkbookDesigner designer = new Aspose.Cells.WorkbookDesigner();
designer.Workbook = new Aspose.Cells.Workbook("Template.xlsx");

// Set the data source for the smart marker
designer.SetDataSource(dt);

// Process the smart markers – the IF condition will evaluate for each row
designer.Process();

// Save the resulting workbook
designer.Workbook.Save("Output.xlsx");
