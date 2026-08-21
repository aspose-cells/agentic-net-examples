// Title: Set Hyperlink Theme Color to Dark Gray Globally in Aspose.Cells for .NET
// Description: Shows how to add a hyperlink, change the workbook's hyperlink theme color to dark gray using SetThemeColor, apply the theme to the default style, and save the file so all hyperlinks render in dark gray.
// Keywords: Aspose.Cells | C# | .NET | hyperlink theme color | dark gray hyperlink | SetThemeColor | default style | global hyperlink color | Excel hyperlink formatting | workbook theme
// Common Searches: Aspose.Cells change hyperlink color .NET | Set hyperlink theme to dark gray in C# | Apply global hyperlink color Aspose.Cells | How to modify hyperlink theme color in Excel using Aspose | C# set workbook theme hyperlink color
// Developer Intent: Modify the workbook’s theme so that every hyperlink appears in dark gray without updating each cell individually.
// Use Cases: Corporate reports that require hyperlinks to match a dark‑gray brand palette. | Multi‑sheet Excel files where all links must retain a consistent color for accessibility. | Retrofitting an existing workbook to improve hyperlink visibility on high‑contrast displays.
// AI Prompts: Generate C# code to set a custom RGB hyperlink theme color in Aspose.Cells after a workbook is opened. | Explain how to verify that existing hyperlinks adopt the new theme color without iterating over each hyperlink. | Provide a step‑by‑step guide to change the hyperlink theme color, save the workbook, and reopen it with the updated color applied.

using Aspose.Cells;
using System.Drawing;

// Create a new workbook
Workbook workbook = new Workbook();
Worksheet sheet = workbook.Worksheets[0];

// Add a hyperlink to cell A1
int hyperlinkIndex = sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");
Hyperlink hyperlink = sheet.Hyperlinks[hyperlinkIndex];
hyperlink.TextToDisplay = "Example Site";

// Set the theme color for hyperlinks to dark gray
workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.DarkGray);

// Ensure the default font style uses the hyperlink theme color
Style defaultStyle = workbook.DefaultStyle;
defaultStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0);
workbook.DefaultStyle = defaultStyle;

// Save the workbook
workbook.Save("HyperlinkThemeDemo.xlsx");
