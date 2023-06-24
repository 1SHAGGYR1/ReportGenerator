namespace ReportGenerator.Models;

public static class DocumentMetrics
{
    public const uint PageSize = 16838;
    public const string TableWidth = "15614";
    public const string TableCellWidth = "7807";
    
    public class Fonts
    {
        public const string HeaderFontSize = "32";
        public const string UnitFontSize = "27";
        public const string SectionFontSize = "22";
        public const string CriterionFontSize = "17";
    }

    public class PageMargins
    {
        public const int Top = 720;
        public const uint Right = 720;
        public const int Bottom = 720;
        public const uint Left = 720;
        public const uint Header = 720;
        public const uint Footer = 720;
    }
}