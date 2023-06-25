namespace ReportGenerator;

public static class DocumentMetrics
{
    public const uint PageSize = 20000;
    public const string TableWidth = "18000";
    public const int TableCellsCount = 3;
    public static string TableCellWidth => (int.Parse(TableWidth) / TableCellsCount).ToString();
    
    public class Fonts
    {
        public const string HeaderFontSize = "37";
        public const string UnitFontSize = "34";
        public const string SectionFontSize = "31";
        public const string CriterionFontSize = "28";
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

    public class TableShadingColors
    {
        public const string Red = "FF0000";
        public const string Orange = "FFA500";
        public const string Yellow = "FFFF00";
        public const string Green = "00FF00";
    }
}