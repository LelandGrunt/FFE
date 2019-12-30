using System;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace FFE
{
    public static class FfeRegEx
    {
        public static string RegExByIndexAndGroup(string input, string pattern, int matchIndex, string group,
                                                  bool ignoreCase = false,
                                                  bool ignorePatternWhitespace = false,
                                                  bool multiline = false,
                                                  bool singleline = true)
        {
            string value = null;

            MatchCollection matches = Regex.Matches(input, pattern, GetRegexOptions(ignoreCase, ignorePatternWhitespace, multiline, singleline));

            if (matches.Count > 0)
            {
                Match match = matches[matchIndex];

                value = match.Success ? match.Groups[group].Value : String.Empty;
            }
            else
            {
                value = String.Empty;
            }

            return value;
        }

        public static RegexOptions WithIgnoreCase(this RegexOptions regexOptions)
        {
            return regexOptions | RegexOptions.IgnoreCase;
        }

        public static RegexOptions WithIgnorePatternWhitespace(this RegexOptions regexOptions)
        {
            return regexOptions | RegexOptions.IgnorePatternWhitespace;
        }

        public static RegexOptions WithMultiline(this RegexOptions regexOptions)
        {
            return regexOptions | RegexOptions.Multiline;
        }

        public static RegexOptions WithSingleline(this RegexOptions regexOptions)
        {
            return regexOptions | RegexOptions.Singleline;
        }

        public static RegexOptions WithCompiled(this RegexOptions regexOptions)
        {
            return regexOptions | RegexOptions.Compiled;
        }

        public static RegexOptions WithExplicitCapture(this RegexOptions regexOptions)
        {
            return regexOptions | RegexOptions.ExplicitCapture;
        }

        #region Private
        private static RegexOptions GetRegexOptions(bool ignoreCase = false,
                                                    bool ignorePatternWhitespace = false,
                                                    bool multiline = false,
                                                    bool singleline = false,
                                                    bool compiled = false,
                                                    bool explictCapture = false)
        {
            RegexOptions options = RegexOptions.None;
            if (ignoreCase)
            {
                options = RegexOptions.IgnoreCase;
            }
            if (ignorePatternWhitespace)
            {
                options = options | RegexOptions.IgnorePatternWhitespace;
            }
            if (multiline)
            {
                options = options | RegexOptions.Multiline;
            }
            if (singleline)
            {
                options = options | RegexOptions.Singleline;
            }
            if (compiled)
            {
                options = options | RegexOptions.Compiled;
            }
            if (explictCapture)
            {
                options = options | RegexOptions.ExplicitCapture;
            }

            return options;
        }
        #endregion
    }

    #region Exceptions
    [Serializable]
    public class RegExException : Exception
    {
        public string Input { get; set; }

        public string Pattern { get; set; }

        public RegExException() { }

        public RegExException(string message) : base(message) { }

        public RegExException(string message, Exception innerException) : base(message, innerException) { }

        protected RegExException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
            Input = info.GetString("Input");
            Pattern = info.GetString("Pattern");
        }

        public override void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            if (info == null)
                throw new ArgumentNullException("info");
            info.AddValue("Input", Input);
            info.AddValue("Pattern", Pattern);
            base.GetObjectData(info, context);
        }
    }
    #endregion
}