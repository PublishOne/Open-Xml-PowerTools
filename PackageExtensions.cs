using System;
using System.IO;
namespace OpenXmlPowerTools;

public static class PackageExtensions
{
    public static Stream GetStream(this IPackagePart part)
    {
        return part.GetStream(FileMode.Open, FileAccess.Read);
    }
}
