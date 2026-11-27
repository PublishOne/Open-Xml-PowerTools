#nullable enable
using DocumentFormat.OpenXml.Packaging;
using OpenXmlRevisions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace OpenXmlPowerTools;

public static class PackageExtensions
{
    public static Stream GetStream(this IPackagePart part)
    {
        return part.GetStream(FileMode.Open, FileAccess.Read);
    }

    public static IPackagePart CreatePart(this IPackage package, Uri partUri, string contentType)
    {
        return package.CreatePart(partUri, contentType, CompressionOption.Normal);
    }

    public static void CreateRelationship(this IPackagePart part, Uri targetUri, TargetMode targetMode, string relationshipType, string id)
    {
        part.Relationships.Create(targetUri, targetMode, relationshipType, id);
    }
}

public class ExecutableRunner
{
    public class RunResults
    {
        public int ExitCode;
        public Exception RunException;
        public StringBuilder Output;
        public StringBuilder Error;
    }

    public static RunResults RunExecutable(string executablePath, string arguments, string workingDirectory)
    {
        var runResults = new RunResults
        {
            Output = new StringBuilder(),
            Error = new StringBuilder(),
            RunException = null
        };
        try
        {
            if (File.Exists(executablePath))
            {
                using (var proc = new Process())
                {
                    proc.StartInfo.FileName = executablePath;
                    proc.StartInfo.Arguments = arguments;
                    proc.StartInfo.WorkingDirectory = workingDirectory;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.OutputDataReceived +=
                        (o, e) => runResults.Output.Append(e.Data).Append(Environment.NewLine);
                    proc.ErrorDataReceived +=
                        (o, e) => runResults.Error.Append(e.Data).Append(Environment.NewLine);
                    proc.Start();
                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();
                    proc.WaitForExit();
                    runResults.ExitCode = proc.ExitCode;
                }
            }
            else
            {
                throw new ArgumentException("Invalid executable path.", nameof(executablePath));
            }
        }
        catch (Exception e)
        {
            runResults.RunException = e;
        }
        return runResults;
    }
}
public static class WordprocessingMLUtil
{
    public static bool GetBoolProp(XElement runProps, XName xName)
    {
        ArgumentNullException.ThrowIfNull(runProps);
        ArgumentNullException.ThrowIfNull(xName);
        var p = runProps.Element(xName);
        if (p == null)
        {
            return false;
        }

        var v = p.Attribute(W.val);
        if (v == null)
        {
            return true;
        }

        var s = v.Value.ToUpperInvariant();
        if (s is "0" or "FALSE")
        {
            return false;
        }

        if (s is "1" or "TRUE")
        {
            return true;
        }

        return false;
    }

    public static int StringToTwips(string twipsOrPoints)
    {
        ArgumentException.ThrowIfNullOrEmpty(twipsOrPoints);
        // if the pos value is in points, not twips
        if (twipsOrPoints.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            decimal decimalValue = decimal.Parse(twipsOrPoints[..^2], CultureInfo.InvariantCulture);
            return (int)(decimalValue * 20);
        }
        return int.Parse(twipsOrPoints, CultureInfo.InvariantCulture);
    }

    public static int? AttributeToTwips(XAttribute attribute)
    {
        if (attribute == null)
        {
            return null;
        }

        string twipsOrPoints = (string)attribute;

        // if the pos value is in points, not twips
        if (twipsOrPoints.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
        {
            decimal decimalValue = decimal.Parse(twipsOrPoints[..^2], CultureInfo.InvariantCulture);
            return (int)(decimalValue * 20);
        }
        if (twipsOrPoints.Contains('.', StringComparison.Ordinal))
        {
            decimal decimalValue = decimal.Parse(twipsOrPoints, CultureInfo.InvariantCulture);
            return (int)decimalValue;
        }
        return int.Parse(twipsOrPoints, CultureInfo.InvariantCulture);
    }

    private static readonly List<XName> AdditionalRunContainerNames = new List<XName>
        {
            W.w + "bdo",
            W.customXml,
            W.dir,
            W.fldSimple,
            W.hyperlink,
            W.moveFrom,
            W.moveTo,
            W.sdtContent
        };

    public static XElement CoalesceAdjacentRunsWithIdenticalFormatting(XElement runContainer)
    {
        ArgumentNullException.ThrowIfNull(runContainer);
        const string dontConsolidate = "DontConsolidate";

        IEnumerable<IGrouping<string, XElement>> groupedAdjacentRunsWithIdenticalFormatting =
            runContainer
                .Elements()
                .GroupAdjacent(ce =>
                {
                    if (ce.Name == W.r)
                    {
                        if (ce.Elements().Count(e => e.Name != W.rPr) != 1)
                        {
                            return dontConsolidate;
                        }

                        if (ce.Attribute(PtOpenXml.AbstractNumId) != null)
                        {
                            return dontConsolidate;
                        }

                        var rPr = ce.Element(W.rPr);
                        string rPrString = rPr != null ? rPr.ToString(SaveOptions.None) : string.Empty;

                        if (ce.Element(W.t) != null)
                        {
                            return "Wt" + rPrString;
                        }

                        if (ce.Element(W.instrText) != null)
                        {
                            return "WinstrText" + rPrString;
                        }

                        return dontConsolidate;
                    }

                    if (ce.Name == W.ins)
                    {
                        if (ce.Elements(W.del).Any())
                        {
                            return dontConsolidate;
#if false
                                // for w:ins/w:del/w:r/w:delText
                                if ((ce.Elements(W.del).Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                                    !ce.Elements().Elements().Elements(W.delText).Any())
                                    return dontConsolidate;

                                XAttribute dateIns = ce.Attribute(W.date);
                                XElement del = ce.Element(W.del);
                                XAttribute dateDel = del.Attribute(W.date);

                                string authorIns = (string) ce.Attribute(W.author) ?? string.Empty;
                                string dateInsString = dateIns != null
                                    ? ((DateTime) dateIns).ToString("s")
                                    : string.Empty;
                                string authorDel = (string) del.Attribute(W.author) ?? string.Empty;
                                string dateDelString = dateDel != null
                                    ? ((DateTime) dateDel).ToString("s")
                                    : string.Empty;

                                return "Wins" +
                                       authorIns +
                                       dateInsString +
                                       authorDel +
                                       dateDelString +
                                       ce.Elements(W.del)
                                           .Elements(W.r)
                                           .Elements(W.rPr)
                                           .Select(rPr => rPr.ToString(SaveOptions.None))
                                           .StringConcatenate();
#endif
                        }

                        // w:ins/w:r/w:t
                        if ((ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1) ||
                            !ce.Elements().Elements(W.t).Any())
                        {
                            return dontConsolidate;
                        }

                        var dateIns2 = ce.Attribute(W.date);

                        string? authorIns2 = (string?)ce.Attribute(W.author) ?? string.Empty;
                        string dateInsString2 = dateIns2 != null
                            ? ((DateTime)dateIns2).ToString("s")
                            : string.Empty;

                        string? idIns2 = (string?)ce.Attribute(W.id) ?? string.Empty;

                        return "Wins2" +
                               authorIns2 +
                               dateInsString2 +
                               idIns2 +
                               ce.Elements()
                                   .Elements(W.rPr)
                                   .Select(rPr => rPr.ToString(SaveOptions.None))
                                   .StringConcatenate();
                    }

                    if (ce.Name == W.del)
                    {
                        if ((ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1) ||
                            !ce.Elements().Elements(W.delText).Any())
                        {
                            return dontConsolidate;
                        }

                        var dateDel2 = ce.Attribute(W.date);

                        string? authorDel2 = (string?)ce.Attribute(W.author) ?? string.Empty;
                        string dateDelString2 = dateDel2 != null ? ((DateTime)dateDel2).ToString("s") : string.Empty;

                        return "Wdel" +
                               authorDel2 +
                               dateDelString2 +
                               ce.Elements(W.r)
                                   .Elements(W.rPr)
                                   .Select(rPr => rPr.ToString(SaveOptions.None))
                                   .StringConcatenate();
                    }

                    return dontConsolidate;
                });

        var runContainerWithConsolidatedRuns = new XElement(runContainer.Name,
            runContainer.Attributes(),
            groupedAdjacentRunsWithIdenticalFormatting.Select(g =>
            {
                if (g.Key == dontConsolidate)
                {
                    return (object)g;
                }

                string textValue = g
                    .Select(r =>
                        r.Descendants()
                            .Where(d => (d.Name == W.t) || (d.Name == W.delText) || (d.Name == W.instrText))
                            .Select(d => d.Value)
                            .StringConcatenate())
                    .StringConcatenate();
                XAttribute? xs = XmlUtil.GetXmlSpaceAttribute(textValue);

                if (g.First().Name == W.r)
                {
                    if (g.First().Element(W.t) != null)
                    {
                        IEnumerable<IEnumerable<XAttribute>> statusAtt =
                            g.Select(r => r.Descendants(W.t).Take(1).Attributes(PtOpenXml.Status));
                        return new XElement(W.r,
                            g.First().Attributes(),
                            g.First().Elements(W.rPr),
                            new XElement(W.t, statusAtt, xs, textValue));
                    }

                    if (g.First().Element(W.instrText) != null)
                    {
                        return new XElement(W.r,
                            g.First().Attributes(),
                            g.First().Elements(W.rPr),
                            new XElement(W.instrText, xs, textValue));
                    }
                }

                if (g.First().Name == W.ins)
                {
                    XElement? firstR = g.First().Element(W.r);
                    return new XElement(W.ins,
                        g.First().Attributes(),
                        new XElement(W.r,
                            firstR?.Attributes(),
                            g.First().Elements(W.r).Elements(W.rPr),
                            new XElement(W.t, xs, textValue)));
                }

                if (g.First().Name == W.del)
                {
                    XElement? firstR = g.First().Element(W.r);
                    return new XElement(W.del,
                        g.First().Attributes(),
                        new XElement(W.r,
                            firstR?.Attributes(),
                            g.First().Elements(W.r).Elements(W.rPr),
                            new XElement(W.delText, xs, textValue)));
                }
                return g;
            }));

        // Process w:txbxContent//w:p
        foreach (XElement txbx in runContainerWithConsolidatedRuns.Descendants(W.txbxContent))
        {
            foreach (XElement txbxPara in txbx.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.p))
            {
                XElement newPara = CoalesceAdjacentRunsWithIdenticalFormatting(txbxPara);
                txbxPara.ReplaceWith(newPara);
            }
        }

        // Process additional run containers.
        var runContainers = runContainerWithConsolidatedRuns
            .Descendants()
            .Where(d => AdditionalRunContainerNames.Contains(d.Name))
            .ToList();
        foreach (XElement container in runContainers)
        {
            XElement newContainer = CoalesceAdjacentRunsWithIdenticalFormatting(container);
            container.ReplaceWith(newContainer);
        }

        return runContainerWithConsolidatedRuns;
    }

    private static readonly Dictionary<XName, int> Order_settings = new Dictionary<XName, int>
        {
            { W.writeProtection, 10},
            { W.view, 20},
            { W.zoom, 30},
            { W.removePersonalInformation, 40},
            { W.removeDateAndTime, 50},
            { W.doNotDisplayPageBoundaries, 60},
            { W.displayBackgroundShape, 70},
            { W.printPostScriptOverText, 80},
            { W.printFractionalCharacterWidth, 90},
            { W.printFormsData, 100},
            { W.embedTrueTypeFonts, 110},
            { W.embedSystemFonts, 120},
            { W.saveSubsetFonts, 130},
            { W.saveFormsData, 140},
            { W.mirrorMargins, 150},
            { W.alignBordersAndEdges, 160},
            { W.bordersDoNotSurroundHeader, 170},
            { W.bordersDoNotSurroundFooter, 180},
            { W.gutterAtTop, 190},
            { W.hideSpellingErrors, 200},
            { W.hideGrammaticalErrors, 210},
            { W.activeWritingStyle, 220},
            { W.proofState, 230},
            { W.formsDesign, 240},
            { W.attachedTemplate, 250},
            { W.linkStyles, 260},
            { W.stylePaneFormatFilter, 270},
            { W.stylePaneSortMethod, 280},
            { W.documentType, 290},
            { W.mailMerge, 300},
            { W.revisionView, 310},
            { W.trackRevisions, 320},
            { W.doNotTrackMoves, 330},
            { W.doNotTrackFormatting, 340},
            { W.documentProtection, 350},
            { W.autoFormatOverride, 360},
            { W.styleLockTheme, 370},
            { W.styleLockQFSet, 380},
            { W.defaultTabStop, 390},
            { W.autoHyphenation, 400},
            { W.consecutiveHyphenLimit, 410},
            { W.hyphenationZone, 420},
            { W.doNotHyphenateCaps, 430},
            { W.showEnvelope, 440},
            { W.summaryLength, 450},
            { W.clickAndTypeStyle, 460},
            { W.defaultTableStyle, 470},
            { W.evenAndOddHeaders, 480},
            { W.bookFoldRevPrinting, 490},
            { W.bookFoldPrinting, 500},
            { W.bookFoldPrintingSheets, 510},
            { W.drawingGridHorizontalSpacing, 520},
            { W.drawingGridVerticalSpacing, 530},
            { W.displayHorizontalDrawingGridEvery, 540},
            { W.displayVerticalDrawingGridEvery, 550},
            { W.doNotUseMarginsForDrawingGridOrigin, 560},
            { W.drawingGridHorizontalOrigin, 570},
            { W.drawingGridVerticalOrigin, 580},
            { W.doNotShadeFormData, 590},
            { W.noPunctuationKerning, 600},
            { W.characterSpacingControl, 610},
            { W.printTwoOnOne, 620},
            { W.strictFirstAndLastChars, 630},
            { W.noLineBreaksAfter, 640},
            { W.noLineBreaksBefore, 650},
            { W.savePreviewPicture, 660},
            { W.doNotValidateAgainstSchema, 670},
            { W.saveInvalidXml, 680},
            { W.ignoreMixedContent, 690},
            { W.alwaysShowPlaceholderText, 700},
            { W.doNotDemarcateInvalidXml, 710},
            { W.saveXmlDataOnly, 720},
            { W.useXSLTWhenSaving, 730},
            { W.saveThroughXslt, 740},
            { W.showXMLTags, 750},
            { W.alwaysMergeEmptyNamespace, 760},
            { W.updateFields, 770},
            { W.footnotePr, 780},
            { W.endnotePr, 790},
            { W.compat, 800},
            { W.docVars, 810},
            { W.rsids, 820},
            { M.mathPr, 830},
            { W.attachedSchema, 840},
            { W.themeFontLang, 850},
            { W.clrSchemeMapping, 860},
            { W.doNotIncludeSubdocsInStats, 870},
            { W.doNotAutoCompressPictures, 880},
            { W.forceUpgrade, 890}, 
            //{W.captions, 900}, 
            { W.readModeInkLockDown, 910},
            { W.smartTagType, 920}, 
            //{W.sl:schemaLibrary, 930}, 
            { W.doNotEmbedSmartTags, 940},
            { W.decimalSymbol, 950},
            { W.listSeparator, 960},
        };

#if false
// from the schema in the standard
        
writeProtection
view
zoom
removePersonalInformation
removeDateAndTime
doNotDisplayPageBoundaries
displayBackgroundShape
printPostScriptOverText
printFractionalCharacterWidth
printFormsData
embedTrueTypeFonts
embedSystemFonts
saveSubsetFonts
saveFormsData
mirrorMargins
alignBordersAndEdges
bordersDoNotSurroundHeader
bordersDoNotSurroundFooter
gutterAtTop
hideSpellingErrors
hideGrammaticalErrors
activeWritingStyle
proofState
formsDesign
attachedTemplate
linkStyles
stylePaneFormatFilter
stylePaneSortMethod
documentType
mailMerge
revisionView
trackRevisions
doNotTrackMoves
doNotTrackFormatting
documentProtection
autoFormatOverride
styleLockTheme
styleLockQFSet
defaultTabStop
autoHyphenation
consecutiveHyphenLimit
hyphenationZone
doNotHyphenateCaps
showEnvelope
summaryLength
clickAndTypeStyle
defaultTableStyle
evenAndOddHeaders
bookFoldRevPrinting
bookFoldPrinting
bookFoldPrintingSheets
drawingGridHorizontalSpacing
drawingGridVerticalSpacing
displayHorizontalDrawingGridEvery
displayVerticalDrawingGridEvery
doNotUseMarginsForDrawingGridOrigin
drawingGridHorizontalOrigin
drawingGridVerticalOrigin
doNotShadeFormData
noPunctuationKerning
characterSpacingControl
printTwoOnOne
strictFirstAndLastChars
noLineBreaksAfter
noLineBreaksBefore
savePreviewPicture
doNotValidateAgainstSchema
saveInvalidXml
ignoreMixedContent
alwaysShowPlaceholderText
doNotDemarcateInvalidXml
saveXmlDataOnly
useXSLTWhenSaving
saveThroughXslt
showXMLTags
alwaysMergeEmptyNamespace
updateFields
footnotePr
endnotePr
compat
docVars
rsids
m:mathPr
attachedSchema
themeFontLang
clrSchemeMapping
doNotIncludeSubdocsInStats
doNotAutoCompressPictures
forceUpgrade
captions
readModeInkLockDown
smartTagType
sl:schemaLibrary
doNotEmbedSmartTags
decimalSymbol
listSeparator
#endif

    private static readonly Dictionary<XName, int> Order_pPr = new Dictionary<XName, int>
        {
            { W.pStyle, 10 },
            { W.keepNext, 20 },
            { W.keepLines, 30 },
            { W.pageBreakBefore, 40 },
            { W.framePr, 50 },
            { W.widowControl, 60 },
            { W.numPr, 70 },
            { W.suppressLineNumbers, 80 },
            { W.pBdr, 90 },
            { W.shd, 100 },
            { W.tabs, 120 },
            { W.suppressAutoHyphens, 130 },
            { W.kinsoku, 140 },
            { W.wordWrap, 150 },
            { W.overflowPunct, 160 },
            { W.topLinePunct, 170 },
            { W.autoSpaceDE, 180 },
            { W.autoSpaceDN, 190 },
            { W.bidi, 200 },
            { W.adjustRightInd, 210 },
            { W.snapToGrid, 220 },
            { W.spacing, 230 },
            { W.ind, 240 },
            { W.contextualSpacing, 250 },
            { W.mirrorIndents, 260 },
            { W.suppressOverlap, 270 },
            { W.jc, 280 },
            { W.textDirection, 290 },
            { W.textAlignment, 300 },
            { W.textboxTightWrap, 310 },
            { W.outlineLvl, 320 },
            { W.divId, 330 },
            { W.cnfStyle, 340 },
            { W.rPr, 350 },
            { W.sectPr, 360 },
            { W.pPrChange, 370 },
        };

    private static readonly Dictionary<XName, int> Order_rPr = new Dictionary<XName, int>
        {
            { W.moveFrom, 5 },
            { W.moveTo, 7 },
            { W.ins, 10 },
            { W.del, 20 },
            { W.rStyle, 30 },
            { W.rFonts, 40 },
            { W.b, 50 },
            { W.bCs, 60 },
            { W.i, 70 },
            { W.iCs, 80 },
            { W.caps, 90 },
            { W.smallCaps, 100 },
            { W.strike, 110 },
            { W.dstrike, 120 },
            { W.outline, 130 },
            { W.shadow, 140 },
            { W.emboss, 150 },
            { W.imprint, 160 },
            { W.noProof, 170 },
            { W.snapToGrid, 180 },
            { W.vanish, 190 },
            { W.webHidden, 200 },
            { W.color, 210 },
            { W.spacing, 220 },
            { W._w, 230 },
            { W.kern, 240 },
            { W.position, 250 },
            { W.sz, 260 },
            { W14.wShadow, 270 },
            { W14.wTextOutline, 280 },
            { W14.wTextFill, 290 },
            { W14.wScene3d, 300 },
            { W14.wProps3d, 310 },
            { W.szCs, 320 },
            { W.highlight, 330 },
            { W.u, 340 },
            { W.effect, 350 },
            { W.bdr, 360 },
            { W.shd, 370 },
            { W.fitText, 380 },
            { W.vertAlign, 390 },
            { W.rtl, 400 },
            { W.cs, 410 },
            { W.em, 420 },
            { W.lang, 430 },
            { W.eastAsianLayout, 440 },
            { W.specVanish, 450 },
            { W.oMath, 460 },
        };

    private static readonly Dictionary<XName, int> Order_tblPr = new Dictionary<XName, int>
        {
            { W.tblStyle, 10 },
            { W.tblpPr, 20 },
            { W.tblOverlap, 30 },
            { W.bidiVisual, 40 },
            { W.tblStyleRowBandSize, 50 },
            { W.tblStyleColBandSize, 60 },
            { W.tblW, 70 },
            { W.jc, 80 },
            { W.tblCellSpacing, 90 },
            { W.tblInd, 100 },
            { W.tblBorders, 110 },
            { W.shd, 120 },
            { W.tblLayout, 130 },
            { W.tblCellMar, 140 },
            { W.tblLook, 150 },
            { W.tblCaption, 160 },
            { W.tblDescription, 170 },
        };

    private static readonly Dictionary<XName, int> Order_tblBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.start, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
        };

    private static readonly Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 },
        };

    private static readonly Dictionary<XName, int> Order_tcBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.start, 20 },
            { W.left, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
            { W.tl2br, 90 },
            { W.tr2bl, 100 },
        };

    private static readonly Dictionary<XName, int> Order_pBdr = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 },
        };

    public static object WmlOrderElementsPerStandard(XNode node)
    {
        if (node is XElement element)
        {
            if (element.Name == W.pPr)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_pPr.TryGetValue(e.Name, out var value1))
                        {
                            return value1;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.rPr)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_rPr.TryGetValue(e.Name, out var value2))
                        {
                            return value2;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.tblPr)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_tblPr.TryGetValue(e.Name, out var value3))
                        {
                            return value3;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.tcPr)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_tcPr.TryGetValue(e.Name, out var value4))
                        {
                            return value4;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.tcBorders)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_tcBorders.TryGetValue(e.Name, out var value5))
                        {
                            return value5;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.tblBorders)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_tblBorders.TryGetValue(e.Name, out var value1))
                        {
                            return value1;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.pBdr)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_pBdr.TryGetValue(e.Name, out var value1))
                        {
                            return value1;
                        }

                        return 999;
                    }));
            }

            if (element.Name == W.p)
            {
                var newP = new XElement(element.Name,
                    element.Attributes(),
                    element.Elements(W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                    element.Elements().Where(e => e.Name != W.pPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));
                return newP;
            }

            if (element.Name == W.r)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements(W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)),
                    element.Elements().Where(e => e.Name != W.rPr).Select(e => (XElement)WmlOrderElementsPerStandard(e)));
            }

            if (element.Name == W.settings)
            {
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Elements().Select(e => (XElement)WmlOrderElementsPerStandard(e)).OrderBy(e =>
                    {
                        if (Order_settings.TryGetValue(e.Name, out var value6))
                        {
                            return value6;
                        }

                        return 999;
                    }));
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => WmlOrderElementsPerStandard(n)));
        }
        return node;
    }
}
public static class XmlUtil
{
    public static XAttribute? GetXmlSpaceAttribute(string? value)
    {
        if (value == null)
        {
            return null;
        }
        if ((value.Length > 0) && ((value[0] == ' ') || (value[^1] == ' ')))
        {
            return new XAttribute(XNamespace.Xml + "space", "preserve");
        }
        else
        {
            return null;
        }
    }

    public static XAttribute? GetXmlSpaceAttribute(char value)
    {
        return value == ' ' ? new XAttribute(XNamespace.Xml + "space", "preserve") : null;
    }
}

public static class Util
{
    public static bool? GetBoolProp(XElement rPr, XName propertyName)
    {
        ArgumentNullException.ThrowIfNull(rPr);
        ArgumentNullException.ThrowIfNull(propertyName);
        XElement? propAtt = rPr.Element(propertyName);
        if (propAtt == null)
        {
            return null;
        }

        XAttribute? val = propAtt.Attribute(W.val);
        if (val == null)
        {
            return true;
        }

        string s = ((string)val).ToUpperInvariant();
        if (s == "1")
        {
            return true;
        }

        if (s == "0")
        {
            return false;
        }

        if (s == "TRUE")
        {
            return true;
        }

        if (s == "FALSE")
        {
            return false;
        }

        if (s == "ON")
        {
            return true;
        }

        if (s == "OFF")
        {
            return false;
        }

        return (bool?)propAtt.Attribute(W.val);
    }
}
public class FileUtils
{
    public static DirectoryInfo GetDateTimeStampedDirectoryInfo(string prefix)
    {
        DateTime now = DateTime.Now;
        string dirName =
            prefix +
            $"-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}";
        return new DirectoryInfo(dirName);
    }

    public static FileInfo GetDateTimeStampedFileInfo(string prefix, string suffix)
    {
        DateTime now = DateTime.Now;
        string fileName =
            prefix +
            $"-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}" +
            suffix;
        return new FileInfo(fileName);
    }

    public static void ThreadSafeCreateDirectory(DirectoryInfo dir)
    {
        while (true)
        {
            if (dir.Exists)
            {
                break;
            }

            try
            {
                dir.Create();
                break;
            }
            catch (IOException)
            {
                System.Threading.Thread.Sleep(50);
            }
        }
    }

    public static void ThreadSafeCopy(FileInfo sourceFile, FileInfo destFile)
    {
        while (true)
        {
            if (destFile.Exists)
            {
                break;
            }

            try
            {
                File.Copy(sourceFile.FullName, destFile.FullName);
                break;
            }
            catch (IOException)
            {
                System.Threading.Thread.Sleep(50);
            }
        }
    }

    public static void ThreadSafeCreateEmptyTextFileIfNotExist(FileInfo file)
    {
        while (true)
        {
            if (file.Exists)
            {
                break;
            }

            try
            {
                File.WriteAllText(file.FullName, "");
                break;
            }
            catch (IOException)
            {
                System.Threading.Thread.Sleep(50);
            }
        }
    }

#if !NET35
    internal static void ThreadSafeAppendAllLines(FileInfo file, string[] strings)
    {
        while (true)
        {
            try
            {
                File.AppendAllLines(file.FullName, strings);
                break;
            }
            catch (IOException)
            {
                System.Threading.Thread.Sleep(50);
            }
        }
    }
#endif

    public static List<string> GetFilesRecursive(DirectoryInfo dir, string searchPattern)
    {
        var fileList = new List<string>();
        GetFilesRecursiveInternal(dir, searchPattern, fileList);
        return fileList;
    }

    private static void GetFilesRecursiveInternal(DirectoryInfo dir, string searchPattern, List<string> fileList)
    {
        fileList.AddRange(dir.GetFiles(searchPattern).Select(file => file.FullName));
        foreach (DirectoryInfo subdir in dir.GetDirectories())
        {
            GetFilesRecursiveInternal(subdir, searchPattern, fileList);
        }
    }

    public static List<string> GetFilesRecursive(DirectoryInfo dir)
    {
        var fileList = new List<string>();
        GetFilesRecursiveInternal(dir, fileList);
        return fileList;
    }

    private static void GetFilesRecursiveInternal(DirectoryInfo dir, List<string> fileList)
    {
        fileList.AddRange(dir.GetFiles().Select(file => file.FullName));
        foreach (DirectoryInfo subdir in dir.GetDirectories())
        {
            GetFilesRecursiveInternal(subdir, fileList);
        }
    }

    public static void CopyStream(Stream source, Stream target)
    {
        const int bufSize = 0x4096;
        var buf = new byte[bufSize];
        int bytesRead;
        while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
        {
            target.Write(buf, 0, bytesRead);
        }
    }
}

public static class PtUtils
{
    public static string SHA1HashStringForUTF8String(string s)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(s);
        byte[] hashBytes = SHA1.HashData(bytes);
        return HexStringFromBytes(hashBytes);
    }

    public static string SHA1HashStringForByteArray(byte[] bytes)
    {
        byte[] hashBytes = SHA1.HashData(bytes);
        return HexStringFromBytes(hashBytes);
    }

    public static string HexStringFromBytes(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        var sb = new StringBuilder();
        foreach (byte b in bytes)
        {
            var hex = b.ToString("x2", CultureInfo.InvariantCulture);
            sb.Append(hex);
        }
        return sb.ToString();
    }

    public static string NormalizeDirName(string dirName)
    {
        string d = dirName.Replace('\\', '/');
        if (d[dirName.Length - 1] is not '/' and not '\\')
        {
            return d + "/";
        }

        return d;
    }

    public static string MakeValidXml(string p)
    {
        return p.Any(c => c < 0x20)
            ? p.Select(c => c < 0x20 ? $"_{(int)c:X}_" : c.ToString()).StringConcatenate()
            : p;
    }

    public static void AddElementIfMissing(XDocument partXDoc, XElement existing, string newElement)
    {
        if (existing != null)
        {
            return;
        }

        var newXElement = XElement.Parse(newElement);
        newXElement.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
        partXDoc.Root?.Add(newXElement);
    }
}
