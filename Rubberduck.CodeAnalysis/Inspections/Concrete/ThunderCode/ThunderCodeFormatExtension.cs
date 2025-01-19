using Rubberduck.Resources.Inspections;
using System.Globalization;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    internal static class ThunderCodeFormatExtension
    {
        public static string ThunderCodeFormat(this string inspectionBase, params object[] args)
        {
            return string.Format(InspectionResults.ResourceManager.GetString("ThunderCode_Base", CultureInfo.CurrentUICulture), string.Format(inspectionBase, args));
        }
    }
}
