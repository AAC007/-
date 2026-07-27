namespace BlankDemandPlanner.Core.Models;

public static class StockCodeNormalizer
{
    public const int OneCNumericCodeLength = 11;

    public static string NormalizeForComparison(string? value)
    {
        var text = (value ?? string.Empty).Trim().ToUpperInvariant();
        if (text.Length == 0)
        {
            return string.Empty;
        }

        return text.All(char.IsDigit)
            ? text.PadLeft(OneCNumericCodeLength, '0')
            : text;
    }
}
