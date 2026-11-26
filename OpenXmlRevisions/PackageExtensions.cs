using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.IO.Packaging;

namespace OpenXmlPowerTools
{
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
}
