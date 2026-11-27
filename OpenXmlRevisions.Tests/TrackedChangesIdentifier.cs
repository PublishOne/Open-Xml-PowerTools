// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenXmlRevisions.Tests;

public class TrackedChangesIdentifier
{
    private readonly Dictionary<Type, bool> _trackedRevisionsElements = [];

    public TrackedChangesIdentifier()
    {
        _trackedRevisionsElements[typeof(CellDeletion)] = true;
        _trackedRevisionsElements[typeof(CellInsertion)] = true;
        _trackedRevisionsElements[typeof(CellMerge)] = true;
        _trackedRevisionsElements[typeof(CustomXmlDelRangeEnd)] = true;
        _trackedRevisionsElements[typeof(CustomXmlDelRangeStart)] = true;
        _trackedRevisionsElements[typeof(CustomXmlInsRangeEnd)] = true;
        _trackedRevisionsElements[typeof(CustomXmlInsRangeStart)] = true;
        _trackedRevisionsElements[typeof(Deleted)] = true;
        _trackedRevisionsElements[typeof(DeletedFieldCode)] = true;
        _trackedRevisionsElements[typeof(DeletedMathControl)] = true;
        _trackedRevisionsElements[typeof(DeletedRun)] = true;
        _trackedRevisionsElements[typeof(DeletedText)] = true;
        _trackedRevisionsElements[typeof(Inserted)] = true;
        _trackedRevisionsElements[typeof(InsertedMathControl)] = true;
        _trackedRevisionsElements[typeof(InsertedMathControl)] = true;
        _trackedRevisionsElements[typeof(InsertedRun)] = true;
        _trackedRevisionsElements[typeof(MoveFrom)] = true;
        _trackedRevisionsElements[typeof(MoveFromRangeEnd)] = true;
        _trackedRevisionsElements[typeof(MoveFromRangeStart)] = true;
        _trackedRevisionsElements[typeof(MoveTo)] = true;
        _trackedRevisionsElements[typeof(MoveToRangeEnd)] = true;
        _trackedRevisionsElements[typeof(MoveToRangeStart)] = true;
        _trackedRevisionsElements[typeof(MoveToRun)] = true;
        _trackedRevisionsElements[typeof(NumberingChange)] = true;
        _trackedRevisionsElements[typeof(ParagraphMarkRunPropertiesChange)] = true;
        _trackedRevisionsElements[typeof(ParagraphPropertiesChange)] = true;
        _trackedRevisionsElements[typeof(RunPropertiesChange)] = true;
        _trackedRevisionsElements[typeof(SectionPropertiesChange)] = true;
        _trackedRevisionsElements[typeof(TableCellPropertiesChange)] = true;
        _trackedRevisionsElements[typeof(TableGridChange)] = true;
        _trackedRevisionsElements[typeof(TablePropertiesChange)] = true;
        _trackedRevisionsElements[typeof(TablePropertyExceptionsChange)] = true;
        _trackedRevisionsElements[typeof(TableRowPropertiesChange)] = true;
    }

    public bool PartHasTrackedRevisions(OpenXmlPart part)
    {
        if (part.RootElement is null)
        {
            return false;
        }
        return part.RootElement.Descendants()
            .Any(e => _trackedRevisionsElements.ContainsKey(e.GetType()));
    }

    public bool DescendantPartsOrSelfHasTrackedRevisions(OpenXmlPart part)
    {
        if (PartHasTrackedRevisions(part))
        {
            return true;
        }
        foreach (var childPart in part.Parts)
        {
            if (PartHasTrackedRevisions(childPart.OpenXmlPart))
            {
                return true;
            }
        }
        return false;
    }

    public ICollection<Type> GetDescendantPartsOrSelfHasTrackedRevisions(OpenXmlPart part)
    {
        HashSet<Type> trackedRevisionTypes = [.. GetPartRevisionTypes(part)];
        foreach (var childPart in part.Parts)
        {
            if (PartHasTrackedRevisions(childPart.OpenXmlPart))
            {
                trackedRevisionTypes.UnionWith(GetPartRevisionTypes(childPart.OpenXmlPart));
            }
        }
        return trackedRevisionTypes;
    }

    public IEnumerable<Type> GetPartRevisionTypes(OpenXmlPart part)
    {
        if (part.RootElement is null)
        {
            return [];
        }
        return part.RootElement.Descendants()
            .Select(e => e.GetType())
            .Where(t => _trackedRevisionsElements.ContainsKey(t))
            .Distinct();
    }
}
