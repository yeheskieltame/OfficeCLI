// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // ==================== Query Layer ====================

    public DocumentNode Get(string path, int depth = 1)
    {
        if (path == "/" || path == "")
            return GetRootNode(depth);

        var parts = ParsePath(path);
        var element = NavigateToElement(parts);
        if (element == null)
            return new DocumentNode { Path = path, Type = "error", Text = $"Path not found: {path}" };

        return ElementToNode(element, path, depth);
    }

    public List<DocumentNode> Query(string selector)
    {
        var results = new List<DocumentNode>();
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return results;

        // Simple selector parser: element[attr=value]
        var parsed = ParseSelector(selector);

        // Determine if main selector targets runs directly (no > parent)
        bool isRunSelector = parsed.ChildSelector == null &&
            (parsed.Element == "r" || parsed.Element == "run");
        bool isPictureSelector = parsed.ChildSelector == null &&
            (parsed.Element == "picture" || parsed.Element == "image" || parsed.Element == "img");
        bool isEquationSelector = parsed.ChildSelector == null &&
            (parsed.Element == "equation" || parsed.Element == "math" || parsed.Element == "formula");

        // Scheme B: generic XML fallback for unrecognized element types
        // Use GenericXmlQuery.ParseSelector which properly handles namespace prefixes (e.g., "a:ln")
        var genericParsed = GenericXmlQuery.ParseSelector(selector);
        bool isKnownType = string.IsNullOrEmpty(genericParsed.element)
            || genericParsed.element is "p" or "paragraph" or "r" or "run"
                or "picture" or "image" or "img"
                or "equation" or "math" or "formula";
        if (!isKnownType && parsed.ChildSelector == null)
        {
            var root = _doc.MainDocumentPart?.Document;
            if (root != null)
                return GenericXmlQuery.Query(root, genericParsed.element, genericParsed.attrs, genericParsed.containsText);
            return results;
        }

        int paraIdx = -1;
        int mathParaIdx = -1;
        foreach (var element in body.ChildElements)
        {
            // Display equations (m:oMathPara) at body level
            if (element.LocalName == "oMathPara" || element is M.Paragraph)
            {
                mathParaIdx++;
                if (isEquationSelector)
                {
                    var latex = FormulaParser.ToLatex(element);
                    if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                    {
                        results.Add(new DocumentNode
                        {
                            Path = $"/body/oMathPara[{mathParaIdx + 1}]",
                            Type = "equation",
                            Text = latex,
                            Format = { ["mode"] = "display" }
                        });
                    }
                }
                continue;
            }

            if (element is Paragraph para)
            {
                paraIdx++;

                if (isEquationSelector)
                {
                    // Check for display equation (oMathPara inside w:p)
                    var oMathParaInPara = para.ChildElements.FirstOrDefault(e => e.LocalName == "oMathPara" || e is M.Paragraph);
                    if (oMathParaInPara != null)
                    {
                        mathParaIdx++;
                        var latex = FormulaParser.ToLatex(oMathParaInPara);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/oMathPara[{mathParaIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "display" }
                            });
                        }
                        continue;
                    }

                    // Find inline math in this paragraph
                    int mathIdx = 0;
                    foreach (var oMath in para.ChildElements.Where(e => e.LocalName == "oMath" || e is M.OfficeMath))
                    {
                        var latex = FormulaParser.ToLatex(oMath);
                        if (parsed.ContainsText == null || latex.Contains(parsed.ContainsText))
                        {
                            results.Add(new DocumentNode
                            {
                                Path = $"/body/p[{paraIdx + 1}]/oMath[{mathIdx + 1}]",
                                Type = "equation",
                                Text = latex,
                                Format = { ["mode"] = "inline" }
                            });
                        }
                        mathIdx++;
                    }
                }
                else if (isPictureSelector)
                {
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        var drawing = run.GetFirstChild<Drawing>();
                        if (drawing != null)
                        {
                            bool noAlt = parsed.Attributes.ContainsKey("__no-alt");
                            if (noAlt)
                            {
                                var docProps = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
                                if (string.IsNullOrEmpty(docProps?.Description?.Value))
                                    results.Add(CreateImageNode(drawing, run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]"));
                            }
                            else
                            {
                                results.Add(CreateImageNode(drawing, run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]"));
                            }
                        }
                        runIdx++;
                    }
                }
                else if (isRunSelector)
                {
                    // Main selector targets runs: search all runs in all paragraphs
                    int runIdx = 0;
                    foreach (var run in GetAllRuns(para))
                    {
                        if (MatchesRunSelector(run, para, parsed))
                        {
                            results.Add(ElementToNode(run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]", 0));
                        }
                        runIdx++;
                    }
                }
                else
                {
                    if (MatchesSelector(para, parsed, paraIdx))
                    {
                        results.Add(ElementToNode(para, $"/body/p[{paraIdx + 1}]", 0));
                    }

                    if (parsed.ChildSelector != null)
                    {
                        int runIdx = 0;
                        foreach (var run in GetAllRuns(para))
                        {
                            if (MatchesRunSelector(run, para, parsed.ChildSelector))
                            {
                                results.Add(ElementToNode(run, $"/body/p[{paraIdx + 1}]/r[{runIdx + 1}]", 0));
                            }
                            runIdx++;
                        }
                    }
                }
            }
        }

        return results;
    }
}
