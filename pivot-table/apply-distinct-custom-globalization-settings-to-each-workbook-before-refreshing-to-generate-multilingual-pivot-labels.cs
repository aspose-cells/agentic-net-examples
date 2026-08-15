// Title: Set Different Globalization Settings per Workbook for Multilingual Pivot Table Labels – Aspose.Cells for .NET (C#)
// Description: Creates two workbooks (English and French), adds sample sales data, builds a pivot table in each workbook, and applies distinct SettableGlobalizationSettings via SettablePivotGlobalizationSettings to customize column, row, and total captions. The pivots are refreshed, calculated, and saved as separate Excel files with localized labels.
// Keywords: Aspose.Cells | C# pivot table globalization | SettableGlobalizationSettings | SettablePivotGlobalizationSettings | multilingual pivot labels | Excel localization | pivot table custom captions | refresh pivot after settings | .NET Excel automation | regional workbook generation
// Common Searches: Aspose.Cells set pivot label language per workbook | customize pivot table captions in C# | multilingual Excel pivot with Aspose.Cells | apply globalization settings to pivot tables | refresh pivot after changing globalization
// Developer Intent: Generate separate workbooks whose pivot tables display language‑specific column, row, and total texts by applying unique globalization settings before refreshing the pivots.
// Use Cases: Produce English and French sales reports with automatically localized pivot headings. | Automate creation of regional Excel workbooks that require distinct label translations without manual editing. | Build multilingual dashboards where each workbook’s pivot table uses its own globalization configuration.
// AI Prompts: Show how to assign different SettableGlobalizationSettings to multiple workbooks and refresh their pivot tables using Aspose.Cells for .NET. | Provide a compact C# example that sets custom column, row, and total texts for a German pivot table and saves the file. | Explain how to programmatically change pivot table label language after workbook creation with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

// Creates two workbooks (English and French), adds sample sales data, builds a pivot table in each workbook, and applies distinct SettableGlobalizationSettings via SettablePivotGlobalizationSettings to customize column, row, and total captions. The pivots are refreshed, calculated, and saved as separate Excel files with localized labels.
class Program
{
    static void Main()
    {
        // ==============================
        // Workbook 1 – English labels
        // ==============================
        Workbook wbEn = new Workbook();                                   // create workbook
        Worksheet wsEn = wbEn.Worksheets[0];

        // sample data
        wsEn.Cells["A1"].PutValue("Product");
        wsEn.Cells["B1"].PutValue("Sales");
        wsEn.Cells["A2"].PutValue("Apple");
        wsEn.Cells["B2"].PutValue(1200);
        wsEn.Cells["A3"].PutValue("Orange");
        wsEn.Cells["B3"].PutValue(800);

        // create pivot table
        int pivotIdxEn = wsEn.PivotTables.Add("A1:B3", "D1", "PivotEn");
        PivotTable ptEn = wsEn.PivotTables[pivotIdxEn];
        ptEn.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        ptEn.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field

        // custom English globalization settings
        SettableGlobalizationSettings gsEn = new SettableGlobalizationSettings();
        SettablePivotGlobalizationSettings pivGsEn = new SettablePivotGlobalizationSettings();
        pivGsEn.SetTextOfColumnLabels("Column Headers");
        pivGsEn.SetTextOfRowLabels("Row Headers");
        pivGsEn.SetTextOfTotal("Total");
        gsEn.PivotSettings = pivGsEn;
        wbEn.Settings.GlobalizationSettings = gsEn;   // apply settings to workbook

        // refresh pivot to apply the texts
        ptEn.RefreshData();
        ptEn.CalculateData();

        // save workbook
        wbEn.Save("Pivot_English.xlsx");

        // ==============================
        // Workbook 2 – French labels
        // ==============================
        Workbook wbFr = new Workbook();                                   // create workbook
        Worksheet wsFr = wbFr.Worksheets[0];

        // sample data (French)
        wsFr.Cells["A1"].PutValue("Produit");
        wsFr.Cells["B1"].PutValue("Ventes");
        wsFr.Cells["A2"].PutValue("Pomme");
        wsFr.Cells["B2"].PutValue(1200);
        wsFr.Cells["A3"].PutValue("Orange");
        wsFr.Cells["B3"].PutValue(800);

        // create pivot table
        int pivotIdxFr = wsFr.PivotTables.Add("A1:B3", "D1", "PivotFr");
        PivotTable ptFr = wsFr.PivotTables[pivotIdxFr];
        ptFr.AddFieldToArea(PivotFieldType.Row, 0);   // Produit as row field
        ptFr.AddFieldToArea(PivotFieldType.Data, 1);  // Ventes as data field

        // custom French globalization settings
        SettableGlobalizationSettings gsFr = new SettableGlobalizationSettings();
        SettablePivotGlobalizationSettings pivGsFr = new SettablePivotGlobalizationSettings();
        pivGsFr.SetTextOfColumnLabels("En-têtes de colonne");
        pivGsFr.SetTextOfRowLabels("En-têtes de ligne");
        pivGsFr.SetTextOfTotal("Total");
        gsFr.PivotSettings = pivGsFr;
        wbFr.Settings.GlobalizationSettings = gsFr;   // apply settings to workbook

        // refresh pivot to apply the texts
        ptFr.RefreshData();
        ptFr.CalculateData();

        // save workbook
        wbFr.Save("Pivot_French.xlsx");
    }
}
