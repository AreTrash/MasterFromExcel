namespace MasterFromExcel
{
    public static class StringExtension
    {
        public static string Trim(this string source, string trimChars)
        {
            return source.Trim(trimChars.ToCharArray());
        }
        
        public static string TrimStart(this string source, string trimChars)
        {
            return source.TrimStart(trimChars.ToCharArray());
        }
        
        public static string TrimEnd(this string source, string trimChars)
        {
            return source.TrimEnd(trimChars.ToCharArray());
        }

        public static string ToTopUpper(this string source)
        {
            if (source.Length == 0) return source;
            return char.ToUpper(source[0]) + source.Substring(1);
        }

        public static bool EqualsWithoutCase(this string source, string other)
        {
            return source.ToLower().Equals(other.ToLower());
        }
    }
}