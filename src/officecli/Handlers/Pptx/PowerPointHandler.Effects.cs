// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// Apply outer shadow effect to ShapeProperties.
    /// Format: "COLOR" or "COLOR-BLUR-ANGLE-DIST" or "COLOR-BLUR-ANGLE-DIST-OPACITY"
    ///   COLOR: hex (e.g. 000000)
    ///   BLUR: blur radius in points, default 4
    ///   ANGLE: direction in degrees, default 45
    ///   DIST: distance in points, default 3
    ///   OPACITY: 0-100 percent, default 40
    /// Examples: "000000", "000000-6-315-4-50", "none"
    /// </summary>
    private static void ApplyShadow(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.OuterShadow>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        var parts = value.Split('-');
        var colorHex = parts[0].TrimStart('#').ToUpperInvariant();
        var blurPt   = parts.Length > 1 ? double.Parse(parts[1]) : 4.0;
        var angleDeg = parts.Length > 2 ? double.Parse(parts[2]) : 45.0;
        var distPt   = parts.Length > 3 ? double.Parse(parts[3]) : 3.0;
        var opacity  = parts.Length > 4 ? double.Parse(parts[4]) : 40.0;

        var shadow = new Drawing.OuterShadow
        {
            BlurRadius    = (long)(blurPt * 12700),
            Distance      = (long)(distPt * 12700),
            Direction     = (int)(angleDeg * 60000),
            Alignment     = Drawing.RectangleAlignmentValues.TopLeft,
            RotateWithShape = false
        };
        var clr = new Drawing.RgbColorModelHex { Val = colorHex };
        clr.AppendChild(new Drawing.Alpha { Val = (int)(opacity * 1000) });
        shadow.AppendChild(clr);
        effectList.AppendChild(shadow);
    }

    /// <summary>
    /// Apply glow effect to ShapeProperties.
    /// Format: "COLOR" or "COLOR-RADIUS" or "COLOR-RADIUS-OPACITY"
    ///   COLOR: hex (e.g. 0070FF)
    ///   RADIUS: glow radius in points, default 8
    ///   OPACITY: 0-100 percent, default 75
    /// Examples: "0070FF", "FF0000-10", "00B0F0-6-60", "none"
    /// </summary>
    private static void ApplyGlow(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.Glow>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        var parts = value.Split('-');
        var colorHex = parts[0].TrimStart('#').ToUpperInvariant();
        var radiusPt = parts.Length > 1 ? double.Parse(parts[1]) : 8.0;
        var opacity  = parts.Length > 2 ? double.Parse(parts[2]) : 75.0;

        var glow = new Drawing.Glow { Radius = (long)(radiusPt * 12700) };
        var clr = new Drawing.RgbColorModelHex { Val = colorHex };
        clr.AppendChild(new Drawing.Alpha { Val = (int)(opacity * 1000) });
        glow.AppendChild(clr);
        effectList.AppendChild(glow);
    }

    /// <summary>
    /// Apply reflection effect to ShapeProperties.
    /// Format: "TYPE" where TYPE is one of:
    ///   tight / small  — tight reflection, touching (stA=52000 endA=300 endPos=55000)
    ///   half           — half reflection (stA=52000 endA=300 endPos=90000)
    ///   full           — full reflection (stA=52000 endA=300 endPos=100000)
    ///   true           — alias for half
    ///   none / false   — remove reflection
    /// </summary>
    private static void ApplyReflection(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.Reflection>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        // endPos controls how much of the shape is reflected
        int endPos = value.ToLowerInvariant() switch
        {
            "tight" or "small" => 55000,
            "true" or "half"   => 90000,
            "full"             => 100000,
            _ => int.TryParse(value, out var pct) ? pct * 1000 : 90000
        };

        var reflection = new Drawing.Reflection
        {
            BlurRadius      = 6350,
            StartOpacity    = 52000,
            StartPosition   = 0,
            EndAlpha        = 300,
            EndPosition     = endPos,
            Distance        = 0,
            Direction       = 5400000,  // 90° — downward
            VerticalRatio   = -100000,  // flip vertically
            Alignment       = Drawing.RectangleAlignmentValues.BottomLeft,
            RotateWithShape = false
        };
        effectList.AppendChild(reflection);
    }
}
