// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Image Helpers ====================

    private static long ParseEmu(string value)
    {
        // Support: raw EMU number, or suffixed with cm/in/pt/px
        value = value.Trim();
        if (value.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 360000);
        if (value.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 914400);
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 12700);
        if (value.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            return (long)(double.Parse(value[..^2]) * 9525);
        return long.Parse(value); // raw EMU
    }

    private static Run CreateImageRun(string relationshipId, long cx, long cy, string altText)
    {
        var inline = new DW.Inline(
            new DW.Extent { Cx = cx, Cy = cy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = (uint)Environment.TickCount, Name = altText, Description = altText },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }
            ),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = altText },
                            new PIC.NonVisualPictureDrawingProperties()
                        ),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                            new A.Stretch(new A.FillRectangle())
                        ),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = cx, Cy = cy }
                            ),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                        )
                    )
                ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        return new Run(new Drawing(inline));
    }

    private static string GetDrawingInfo(Drawing drawing)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var parts = new List<string>();
        if (docProps?.Description?.Value is string desc && !string.IsNullOrEmpty(desc))
            parts.Add($"alt=\"{desc}\"");
        else if (docProps?.Name?.Value is string name && !string.IsNullOrEmpty(name))
            parts.Add($"name=\"{name}\"");
        if (extent != null)
        {
            var wCm = extent.Cx != null ? $"{extent.Cx.Value / 360000.0:F1}cm" : "?";
            var hCm = extent.Cy != null ? $"{extent.Cy.Value / 360000.0:F1}cm" : "?";
            parts.Add($"{wCm}×{hCm}");
        }
        return parts.Count > 0 ? string.Join(", ", parts) : "unknown";
    }

    private static DocumentNode CreateImageNode(Drawing drawing, Run run, string path)
    {
        var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
        var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();

        var node = new DocumentNode
        {
            Path = path,
            Type = "picture",
            Text = docProps?.Description?.Value ?? docProps?.Name?.Value ?? ""
        };
        if (extent?.Cx != null) node.Format["width"] = $"{extent.Cx.Value / 360000.0:F1}cm";
        if (extent?.Cy != null) node.Format["height"] = $"{extent.Cy.Value / 360000.0:F1}cm";
        if (docProps?.Description?.Value != null) node.Format["alt"] = docProps.Description.Value;

        return node;
    }
}
