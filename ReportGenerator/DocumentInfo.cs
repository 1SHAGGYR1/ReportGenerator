namespace ReportGenerator;

public class DocumentInfo
{
    public ChildInfo ChildInfo { get; set; } = new();

    public DateOnly PeriodStartDate { get; set; }
    
    public DateOnly PeriodEndDate { get; set; }
}