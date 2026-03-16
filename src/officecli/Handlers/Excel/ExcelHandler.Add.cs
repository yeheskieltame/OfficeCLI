// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public string Add(string parentPath, string type, int? index, Dictionary<string, string> properties)
    {
        switch (type.ToLowerInvariant())
        {
            case "sheet":
                var workbookPart = _doc.WorkbookPart
                    ?? throw new InvalidOperationException("Workbook not found");
                var sheets = GetWorkbook().GetFirstChild<Sheets>()
                    ?? GetWorkbook().AppendChild(new Sheets());

                var name = properties.GetValueOrDefault("name", $"Sheet{sheets.Elements<Sheet>().Count() + 1}");
                var newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                newWorksheetPart.Worksheet.Save();

                var sheetId = sheets.Elements<Sheet>().Any()
                    ? sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1
                    : 1;
                var relId = workbookPart.GetIdOfPart(newWorksheetPart);

                sheets.AppendChild(new Sheet { Id = relId, SheetId = (uint)sheetId, Name = name });
                GetWorkbook().Save();
                return $"/{name}";

            case "row":
                var segments = parentPath.TrimStart('/').Split('/', 2);
                var sheetName = segments[0];
                var worksheet = FindWorksheet(sheetName)
                    ?? throw new ArgumentException($"Sheet not found: {sheetName}");
                var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(worksheet).AppendChild(new SheetData());

                var rowIdx = index ?? ((int)(sheetData.Elements<Row>().LastOrDefault()?.RowIndex?.Value ?? 0) + 1);
                var newRow = new Row { RowIndex = (uint)rowIdx };

                // Create cells if cols specified
                if (properties.TryGetValue("cols", out var colsStr))
                {
                    var cols = int.Parse(colsStr);
                    for (int c = 0; c < cols; c++)
                    {
                        var colLetter = IndexToColumnName(c + 1);
                        newRow.AppendChild(new Cell { CellReference = $"{colLetter}{rowIdx}" });
                    }
                }

                var afterRow = sheetData.Elements<Row>().LastOrDefault(r => (r.RowIndex?.Value ?? 0) < (uint)rowIdx);
                if (afterRow != null)
                    afterRow.InsertAfterSelf(newRow);
                else
                    sheetData.InsertAt(newRow, 0);

                GetSheet(worksheet).Save();
                return $"/{sheetName}/row[{rowIdx}]";

            case "cell":
                var cellSegments = parentPath.TrimStart('/').Split('/', 2);
                var cellSheetName = cellSegments[0];
                var cellWorksheet = FindWorksheet(cellSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cellSheetName}");
                var cellSheetData = GetSheet(cellWorksheet).GetFirstChild<SheetData>()
                    ?? GetSheet(cellWorksheet).AppendChild(new SheetData());

                var cellRef = properties.GetValueOrDefault("ref", "A1");
                var cell = FindOrCreateCell(cellSheetData, cellRef);

                if (properties.TryGetValue("value", out var value))
                {
                    cell.CellValue = new CellValue(value);
                    if (!double.TryParse(value, out _))
                        cell.DataType = new EnumValue<CellValues>(CellValues.String);
                }
                if (properties.TryGetValue("formula", out var formula))
                {
                    cell.CellFormula = new CellFormula(formula);
                    cell.CellValue = null;
                }
                if (properties.TryGetValue("type", out var cellType))
                {
                    cell.DataType = cellType.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        _ => cell.DataType
                    };
                }
                if (properties.TryGetValue("clear", out _))
                {
                    cell.CellValue = null;
                    cell.CellFormula = null;
                }

                // Apply style properties if any
                var cellStyleProps = new Dictionary<string, string>();
                foreach (var (key, val) in properties)
                {
                    if (ExcelStyleManager.IsStyleKey(key))
                        cellStyleProps[key] = val;
                }
                if (cellStyleProps.Count > 0)
                {
                    var cellWbPart = _doc.WorkbookPart
                        ?? throw new InvalidOperationException("Workbook not found");
                    var styleManager = new ExcelStyleManager(cellWbPart);
                    cell.StyleIndex = styleManager.ApplyStyle(cell, cellStyleProps);
                }

                GetSheet(cellWorksheet).Save();
                return $"/{cellSheetName}/{cellRef}";

            case "databar":
            case "conditionalformatting":
            {
                var cfSegments = parentPath.TrimStart('/').Split('/', 2);
                var cfSheetName = cfSegments[0];
                var cfWorksheet = FindWorksheet(cfSheetName)
                    ?? throw new ArgumentException($"Sheet not found: {cfSheetName}");

                var sqref = properties.GetValueOrDefault("sqref", "A1:A10");
                var minVal = properties.GetValueOrDefault("min", "0");
                var maxVal = properties.GetValueOrDefault("max", "1");
                var cfColor = properties.GetValueOrDefault("color", "638EC6");
                var normalizedColor = (cfColor.Length == 6 ? "FF" : "") + cfColor.ToUpperInvariant();

                var cfRule = new ConditionalFormattingRule
                {
                    Type = ConditionalFormatValues.DataBar,
                    Priority = 1
                };
                var dataBar = new DataBar();
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Number,
                    Val = minVal
                });
                dataBar.Append(new ConditionalFormatValueObject
                {
                    Type = ConditionalFormatValueObjectValues.Number,
                    Val = maxVal
                });
                dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = normalizedColor });
                cfRule.Append(dataBar);

                var cf = new ConditionalFormatting(cfRule)
                {
                    SequenceOfReferences = new ListValue<StringValue>(
                        sqref.Split(' ').Select(s => new StringValue(s)))
                };

                // Insert after sheetData (or after existing elements)
                var wsElement = GetSheet(cfWorksheet);
                var sheetDataEl = wsElement.GetFirstChild<SheetData>();
                if (sheetDataEl != null)
                    sheetDataEl.InsertAfterSelf(cf);
                else
                    wsElement.Append(cf);

                GetSheet(cfWorksheet).Save();
                return $"/{cfSheetName}/conditionalFormatting[{sqref}]";
            }

            default:
            {
                // Generic fallback: create typed element via SDK schema validation
                // Parse parentPath: /<SheetName>/xmlPath...
                var fbSegments = parentPath.TrimStart('/').Split('/', 2);
                var fbSheetName = fbSegments[0];
                var fbWorksheet = FindWorksheet(fbSheetName);
                if (fbWorksheet == null)
                    throw new ArgumentException($"Sheet not found: {fbSheetName}");

                OpenXmlElement fbParent = GetSheet(fbWorksheet);
                if (fbSegments.Length > 1 && !string.IsNullOrEmpty(fbSegments[1]))
                {
                    var xmlSegments = GenericXmlQuery.ParsePathSegments(fbSegments[1]);
                    fbParent = GenericXmlQuery.NavigateByPath(fbParent!, xmlSegments)
                        ?? throw new ArgumentException($"Parent element not found: {parentPath}");
                }

                var created = GenericXmlQuery.TryCreateTypedElement(fbParent!, type, properties, index);
                if (created == null)
                    throw new ArgumentException($"Schema-invalid element type '{type}' for parent '{parentPath}'. " +
                        "Use raw-set --action append with explicit XML instead.");

                GetSheet(fbWorksheet).Save();

                var siblings = fbParent.ChildElements.Where(e => e.LocalName == created.LocalName).ToList();
                var createdIdx = siblings.IndexOf(created) + 1;
                return $"{parentPath}/{created.LocalName}[{createdIdx}]";
            }
        }
    }

    public void Remove(string path)
    {
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        if (segments.Length == 1)
        {
            // Remove entire sheet
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var sheets = GetWorkbook().GetFirstChild<Sheets>();
            var sheet = sheets?.Elements<Sheet>()
                .FirstOrDefault(s => s.Name?.Value?.Equals(sheetName, StringComparison.OrdinalIgnoreCase) == true);
            if (sheet == null)
                throw new ArgumentException($"Sheet not found: {sheetName}");

            var relId = sheet.Id?.Value;
            sheet.Remove();
            if (relId != null)
                workbookPart.DeletePart(workbookPart.GetPartById(relId));
            GetWorkbook().Save();
            return;
        }

        // Remove cell or row
        var cellRef = segments[1];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Check if it's a row reference like row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            row.Remove();
        }
        else
        {
            // Cell reference
            var cell = FindCell(sheetData, cellRef)
                ?? throw new ArgumentException($"Cell {cellRef} not found");
            cell.Remove();
        }

        GetSheet(worksheet).Save();
    }

    public string Move(string sourcePath, string? targetParentPath, int? index)
    {
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot move an entire sheet. Use move on rows or elements within a sheet.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Determine target
        string effectiveParentPath;
        SheetData targetSheetData;
        if (string.IsNullOrEmpty(targetParentPath))
        {
            effectiveParentPath = $"/{sheetName}";
            targetSheetData = sheetData;
        }
        else
        {
            effectiveParentPath = targetParentPath;
            var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
            var tgtWorksheet = FindWorksheet(tgtSegments[0])
                ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
            targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
                ?? throw new ArgumentException("Target sheet has no data");
        }

        // Find and move the row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            row.Remove();

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(row);
                else
                    targetSheetData.AppendChild(row);
            }
            else
            {
                targetSheetData.AppendChild(row);
            }

            GetSheet(worksheet).Save();
            var newRows = targetSheetData.Elements<Row>().ToList();
            var newIdx = newRows.IndexOf(row) + 1;
            return $"{effectiveParentPath}/row[{newIdx}]";
        }

        throw new ArgumentException($"Move not supported for: {elementRef}. Supported: row[N]");
    }

    public string CopyFrom(string sourcePath, string targetParentPath, int? index)
    {
        var segments = sourcePath.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];
        var worksheet = FindWorksheet(sheetName)
            ?? throw new ArgumentException($"Sheet not found: {sheetName}");

        if (segments.Length < 2)
            throw new ArgumentException("Cannot copy an entire sheet with --from. Use add --type sheet instead.");

        var elementRef = segments[1];
        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Sheet has no data");

        // Find target
        var tgtSegments = targetParentPath.TrimStart('/').Split('/', 2);
        var tgtWorksheet = FindWorksheet(tgtSegments[0])
            ?? throw new ArgumentException($"Target sheet not found: {tgtSegments[0]}");
        var targetSheetData = GetSheet(tgtWorksheet).GetFirstChild<SheetData>()
            ?? throw new ArgumentException("Target sheet has no data");

        // Copy row
        var rowMatch = Regex.Match(elementRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIdx)
                ?? throw new ArgumentException($"Row {rowIdx} not found");
            var clone = (Row)row.CloneNode(true);

            if (index.HasValue)
            {
                var rows = targetSheetData.Elements<Row>().ToList();
                if (index.Value >= 0 && index.Value < rows.Count)
                    rows[index.Value].InsertBeforeSelf(clone);
                else
                    targetSheetData.AppendChild(clone);
            }
            else
            {
                targetSheetData.AppendChild(clone);
            }

            GetSheet(tgtWorksheet).Save();
            var newRows = targetSheetData.Elements<Row>().ToList();
            var newIdx = newRows.IndexOf(clone) + 1;
            return $"{targetParentPath}/row[{newIdx}]";
        }

        throw new ArgumentException($"Copy not supported for: {elementRef}. Supported: row[N]");
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var workbookPart = _doc.WorkbookPart
            ?? throw new InvalidOperationException("No workbook part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a worksheet's DrawingsPart
                var sheetName = parentPartPath.TrimStart('/');
                var worksheetPart = FindWorksheet(sheetName)
                    ?? throw new ArgumentException(
                        $"Sheet not found: {sheetName}. Chart must be added under a sheet: add-part <file> /<SheetName> --type chart");

                var drawingsPart = worksheetPart.DrawingsPart
                    ?? worksheetPart.AddNewPart<DrawingsPart>();

                // Initialize DrawingsPart if new
                if (drawingsPart.WorksheetDrawing == null)
                {
                    drawingsPart.WorksheetDrawing =
                        new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
                    drawingsPart.WorksheetDrawing.Save();

                    // Link DrawingsPart to worksheet if not already linked
                    if (GetSheet(worksheetPart).GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>() == null)
                    {
                        var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
                        GetSheet(worksheetPart).Append(
                            new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = drawingRelId });
                        GetSheet(worksheetPart).Save();
                    }
                }

                var chartPart = drawingsPart.AddNewPart<ChartPart>();
                var relId = drawingsPart.GetIdOfPart(chartPart);

                // Initialize with minimal valid ChartSpace
                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = drawingsPart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/{sheetName}/chart[{chartIdx + 1}]");

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart");
        }
    }
}
