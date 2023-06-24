namespace Shared;

public class Unit
{
    public int Digit { get; set; }

    public string Text { get; set; }

    public List<Section> SectionList { get; set; } = new();

    public List<Criterion> UnSectionedCriterionList { get; set; } = new();
}