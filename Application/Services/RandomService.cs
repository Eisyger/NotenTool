namespace Application.Services;

public class RandomService
{
    public List<T> Shuffle<T>(List<T> list, string seed)
    {
        var random = new Random(GetStableSeed(seed));

        var result = new List<T>(list);

        for (var i = result.Count - 1; i > 0; i--)
        {
            var j = random.Next(i + 1);
            (result[i], result[j]) = (result[j], result[i]);
        }

        return result;
    }

    private static int GetStableSeed(string text)
    {
        unchecked
        {
            const int offset = unchecked((int)2166136261);
            const int prime = 16777619;

            var hash = offset;

            foreach (var c in text)
            {
                hash ^= c;
                hash *= prime;
            }

            return hash & int.MaxValue;
        }
    }
}