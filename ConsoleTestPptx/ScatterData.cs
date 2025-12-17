namespace ConsoleTestPptx;

public class ScatterData
{
	public string Name { get; set; } = string.Empty;
	public List<KeyValuePair<float, float>> DataPoints { get; set; } = new();
}