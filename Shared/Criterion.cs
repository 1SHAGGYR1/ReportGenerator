namespace Shared;

public class Criterion
{
    public int Key { get; set; }
    
    public string Text { get; set; }
    
    public SortedDictionary<int, Criterion> Inner { get; set; } = new();
}