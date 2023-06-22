namespace MyDocumentParses;

public class Criteria
{
    public string Key { get; set; }
    
    public string Text { get; set; }

    
    public Dictionary<string, Criteria> Inner { get; set; }
}