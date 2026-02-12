namespace Application.Services;

public class RandomService
{
    public List<T> Shuffle<T>(List<T> list, string seed)
    {
        var random = new Random(GetPositiveHash(seed));
        return list.OrderBy(x => random.Next()).ToList();
    }
    
    private static int GetPositiveHash(string text)
    {
        return Math.Abs(text.GetHashCode());
    }
}


