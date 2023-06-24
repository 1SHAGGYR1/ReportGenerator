namespace Shared;

public class Section
{
    public int StartCriterionKey { get; set; }
    
    public int EndCriterionKey { get; set; }

    public string Text { get; set; }

    public List<Criterion> CriterionList { get; set; } = new();
}