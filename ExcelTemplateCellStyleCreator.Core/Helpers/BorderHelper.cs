namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Converts between boolean border flags and the "lrtb" selection string format.
    /// </summary>
    public static class BorderHelper
    {
        /// <summary>
        /// Builds a selection string from individual border flags (e.g. "lrt" for left+right+top).
        /// </summary>
        public static string BuildSelectionString(bool left, bool right, bool top, bool bottom)
        {
            var sb = new System.Text.StringBuilder(4);
            if (left)   sb.Append('l');
            if (right)  sb.Append('r');
            if (top)    sb.Append('t');
            if (bottom) sb.Append('b');
            return sb.ToString();
        }

        /// <summary>
        /// Parses a selection string into individual border flags.
        /// </summary>
        public static (bool Left, bool Right, bool Top, bool Bottom) ParseSelection(string selection)
        {
            if (string.IsNullOrEmpty(selection) || selection == "(none)")
                return (false, false, false, false);

            return (selection.Contains('l'), selection.Contains('r'),
                    selection.Contains('t'), selection.Contains('b'));
        }

        /// <summary>
        /// Returns the selection string for display, showing "(none)" if empty.
        /// </summary>
        public static string FormatForDisplay(string selection)
            => string.IsNullOrEmpty(selection) ? "(none)" : selection;
    }
}
