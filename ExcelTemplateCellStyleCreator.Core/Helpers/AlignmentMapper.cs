using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelTemplateCellStyleCreator.Core
{
    /// <summary>
    /// Maps between user-friendly alignment labels ("Left", "Center", "Right", "Top", "Bottom")
    /// and OpenXML alignment enum values.
    /// </summary>
    public static class AlignmentMapper
    {
        public static HorizontalAlignmentValues ToHorizontalEnum(string? label)
        {
            if (label == "Center") return HorizontalAlignmentValues.Center;
            if (label == "Right") return HorizontalAlignmentValues.Right;
            return HorizontalAlignmentValues.Left;
        }

        public static VerticalAlignmentValues ToVerticalEnum(string? label)
        {
            if (label == "Top") return VerticalAlignmentValues.Top;
            if (label == "Bottom") return VerticalAlignmentValues.Bottom;
            return VerticalAlignmentValues.Center;
        }

        public static string ToHorizontalLabel(HorizontalAlignmentValues? value)
        {
            if (value == null) return "Left";
            if (value == HorizontalAlignmentValues.Center) return "Center";
            if (value == HorizontalAlignmentValues.Right) return "Right";
            return "Left";
        }

        public static string ToVerticalLabel(VerticalAlignmentValues? value)
        {
            if (value == null) return "Center";
            if (value == VerticalAlignmentValues.Top) return "Top";
            if (value == VerticalAlignmentValues.Bottom) return "Bottom";
            return "Center";
        }
    }
}
