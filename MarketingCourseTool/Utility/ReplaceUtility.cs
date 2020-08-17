using System.Text.RegularExpressions;

namespace MarketingCourseTool.Utility
{
    public static class ReplaceUtility
    {
        public static string Replace(string baseString, string replaceTarget, string replaceValue)
        {
            if (string.IsNullOrEmpty(replaceValue)) return baseString;
            return Regex.Replace(baseString, replaceTarget, replaceValue);
        }
    }
}
