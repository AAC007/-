using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using BlankDemandPlanner.Core.Enums;
using BlankDemandPlanner.Core.Interfaces;
using BlankDemandPlanner.Core.Models;

namespace BlankDemandPlanner.Services.Normalization;

public sealed partial class BlankNormalizationService : IBlankNormalizationService
{
    public Task<BlankNormalizationResult> NormalizeAsync(string sourceName, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var warnings = new List<string>();
        var normalizedSource = NormalizeSpaces(sourceName).ToUpperInvariant();
        var type = DetectType(normalizedSource);
        var material = DetectMaterial(normalizedSource);
        var gosts = GostRegex().Matches(normalizedSource).Select(m => m.Value.Replace("ГОСТ ", "ГОСТ ", StringComparison.Ordinal)).Distinct().ToList();
        var profileGost = gosts.FirstOrDefault(g => IsProfileGost(g));
        var materialGost = gosts.FirstOrDefault(g => !IsProfileGost(g));

        var diameter = ReadDecimal(DiameterRegex().Match(normalizedSource));
        var length = ReadDecimal(LengthRegex().Match(normalizedSource));
        var thickness = ReadDecimal(ThicknessRegex().Match(normalizedSource));
        var wall = ReadDecimal(WallRegex().Match(normalizedSource));
        var (width, height, dimensionLength) = DetectDimensions(normalizedSource);

        if (type is BlankType.RoundBar or BlankType.PipeRound && diameter is null)
        {
            diameter = ReadLeadingSizeAfterKeyword(normalizedSource, "КРУГ");
        }

        if (type is BlankType.SquareBar && width is null)
        {
            width = ReadLeadingSizeAfterKeyword(normalizedSource, "КВАДРАТ");
            height = width;
        }

        if (length is null)
        {
            length = dimensionLength;
        }

        if (width is not null && height is not null && length is not null &&
            type is not BlankType.PipeRectangular and not BlankType.PipeRound)
        {
            type = BlankType.Sheet;
        }

        if (material is null)
        {
            warnings.Add("Материал не распознан");
        }

        if (type == BlankType.Unknown)
        {
            warnings.Add("Тип заготовки не распознан");
        }

        var normalizedName = BuildNormalizedName(type, material, materialGost, profileGost, diameter, width, height, thickness, wall, length);
        var canonicalKey = BuildCanonicalKey(type, material, diameter, width, height, thickness, wall, length);
        var confidence = CalculateConfidence(type, material, profileGost, diameter, width, height, thickness);

        return Task.FromResult(new BlankNormalizationResult(
            type,
            material,
            materialGost,
            profileGost,
            diameter,
            width,
            height,
            thickness,
            wall,
            length,
            normalizedName,
            canonicalKey,
            confidence,
            warnings));
    }

    private static BlankType DetectType(string value)
    {
        if (value.Contains("БРОНЗ", StringComparison.Ordinal) && value.Contains("ЛИСТ", StringComparison.Ordinal)) return BlankType.BronzeSheet;
        if (value.Contains("БРОНЗ", StringComparison.Ordinal) && (value.Contains("ПРУТОК", StringComparison.Ordinal) || value.Contains("КРУГ", StringComparison.Ordinal))) return BlankType.BronzeBar;
        if (value.Contains("ШЕСТИГР", StringComparison.Ordinal)) return BlankType.HexBar;
        if (value.Contains("КВАДРАТ", StringComparison.Ordinal)) return BlankType.SquareBar;
        if (value.Contains("КРУГ", StringComparison.Ordinal) || value.Contains("D", StringComparison.Ordinal) && !value.Contains("ТРУБ", StringComparison.Ordinal)) return BlankType.RoundBar;
        if (value.Contains("ТРУБ", StringComparison.Ordinal) && (value.Contains("ПРЯМОУГ", StringComparison.Ordinal) || RectPipeRegex().IsMatch(value))) return BlankType.PipeRectangular;
        if (value.Contains("ТРУБ", StringComparison.Ordinal)) return BlankType.PipeRound;
        if (value.Contains("УГОЛ", StringComparison.Ordinal)) return BlankType.Angle;
        if (value.Contains("ШВЕЛЛ", StringComparison.Ordinal)) return BlankType.Channel;
        if (value.Contains("ДВУТАВ", StringComparison.Ordinal)) return BlankType.IBeam;
        if (value.Contains("СВАР", StringComparison.Ordinal)) return BlankType.WeldingElement;
        if (value.Contains("ПОКОВ", StringComparison.Ordinal)) return BlankType.Forging;
        if (value.Contains("ПОКУП", StringComparison.Ordinal)) return BlankType.Purchased;
        if (value.Contains("ЛИТ", StringComparison.Ordinal)) return BlankType.Casting;
        if (value.Contains("ПЛИТ", StringComparison.Ordinal)) return BlankType.Plate;
        if (value.Contains("ЛИСТ", StringComparison.Ordinal)) return BlankType.Sheet;
        if (value.Contains("ЗАГОТОВ", StringComparison.Ordinal)) return BlankType.CustomBlank;
        return BlankType.Unknown;
    }

    private static string? DetectMaterial(string value)
    {
        var match = MaterialRegex().Match(value);
        if (!match.Success)
        {
            return null;
        }

        var prefix = match.Groups["prefix"].Value;
        var grade = match.Groups["grade"].Value.Replace(" ", string.Empty, StringComparison.Ordinal);
        return string.IsNullOrWhiteSpace(prefix) ? grade : $"{ToTitle(prefix)} {grade}";
    }

    private static (decimal? width, decimal? height, decimal? length) DetectDimensions(string value)
    {
        var match = DimensionsRegex().Match(value);
        if (!match.Success)
        {
            return (null, null, null);
        }

        var values = match.Groups["n"].Captures.Select(c => ParseDecimal(c.Value)).ToList();
        return values.Count switch
        {
            2 => (values[0], values[1], null),
            >= 3 => (values[0], values[1], values[2]),
            _ => (null, null, null)
        };
    }

    private static decimal? ReadDecimal(Match match) =>
        match.Success ? ParseDecimal(match.Groups[1].Value) : null;

    private static decimal? ReadLeadingSizeAfterKeyword(string value, string keyword)
    {
        var index = value.IndexOf(keyword, StringComparison.Ordinal);
        if (index < 0)
        {
            return null;
        }

        var match = NumberRegex().Match(value[(index + keyword.Length)..]);
        return match.Success ? ParseDecimal(match.Value) : null;
    }

    private static decimal ParseDecimal(string value) =>
        decimal.Parse(value.Replace(",", ".", StringComparison.Ordinal), NumberStyles.Number, CultureInfo.InvariantCulture);

    private static string BuildNormalizedName(BlankType type, string? material, string? materialGost, string? profileGost, decimal? diameter, decimal? width, decimal? height, decimal? thickness, decimal? wall, decimal? length)
    {
        var builder = new StringBuilder(DisplayName(type));
        if (diameter is not null) builder.Append(" D").Append(Format(diameter));
        if (width is not null && height is not null) builder.Append(' ').Append(Format(width)).Append('x').Append(Format(height));
        else if (width is not null) builder.Append(' ').Append(Format(width));
        if (thickness is not null) builder.Append(" S").Append(Format(thickness));
        if (wall is not null) builder.Append(" Wall").Append(Format(wall));
        if (length is not null) builder.Append(" L").Append(Format(length));
        if (!string.IsNullOrWhiteSpace(material)) builder.Append(' ').Append(material);
        if (!string.IsNullOrWhiteSpace(profileGost)) builder.Append(' ').Append(profileGost);
        if (!string.IsNullOrWhiteSpace(materialGost)) builder.Append(" / ").Append(materialGost);
        return builder.ToString();
    }

    private static string BuildCanonicalKey(BlankType type, string? material, decimal? diameter, decimal? width, decimal? height, decimal? thickness, decimal? wall, decimal? length)
    {
        var parts = new List<string> { type.ToString().ToUpperInvariant(), NormalizeMaterialKey(material) };
        if (diameter is not null) parts.Add($"D{Format(diameter)}");
        if (width is not null) parts.Add($"W{Format(width)}");
        if (height is not null) parts.Add($"H{Format(height)}");
        if (thickness is not null) parts.Add($"S{Format(thickness)}");
        if (wall is not null) parts.Add($"T{Format(wall)}");
        if (length is not null) parts.Add($"L{Format(length)}");
        return string.Join('|', parts.Where(p => !string.IsNullOrWhiteSpace(p)));
    }

    private static double CalculateConfidence(BlankType type, string? material, string? profileGost, decimal? diameter, decimal? width, decimal? height, decimal? thickness)
    {
        var score = 0.15;
        if (type != BlankType.Unknown) score += 0.25;
        if (!string.IsNullOrWhiteSpace(material)) score += 0.25;
        if (!string.IsNullOrWhiteSpace(profileGost)) score += 0.1;
        if (diameter is not null || width is not null || height is not null || thickness is not null) score += 0.25;
        return Math.Min(1, score);
    }

    private static bool IsProfileGost(string gost) =>
        gost.Contains("2590", StringComparison.Ordinal) ||
        gost.Contains("2591", StringComparison.Ordinal) ||
        gost.Contains("2879", StringComparison.Ordinal) ||
        gost.Contains("19903", StringComparison.Ordinal) ||
        gost.Contains("8732", StringComparison.Ordinal) ||
        gost.Contains("8645", StringComparison.Ordinal);

    private static string NormalizeSpaces(string value) => SpacesRegex().Replace(value.Trim(), " ");

    private static string NormalizeMaterialKey(string? material) =>
        string.IsNullOrWhiteSpace(material)
            ? string.Empty
            : material.ToUpperInvariant().Replace("СТАЛЬ", string.Empty, StringComparison.Ordinal).Replace("НЕРЖ", string.Empty, StringComparison.Ordinal).Replace(" ", string.Empty, StringComparison.Ordinal);

    private static string DisplayName(BlankType type) => type switch
    {
        BlankType.RoundBar => "Круг",
        BlankType.SquareBar => "Квадрат",
        BlankType.HexBar => "Шестигранник",
        BlankType.Sheet => "Лист",
        BlankType.Plate => "Плита",
        BlankType.PipeRound => "Труба профильная круглая",
        BlankType.PipeRectangular => "Труба профильная прямоугольная",
        BlankType.Angle => "Уголок",
        BlankType.Channel => "Швеллер",
        BlankType.IBeam => "Двутавр",
        BlankType.BronzeBar => "Пруток бронзовый",
        BlankType.BronzeSheet => "Лист бронзовый",
        BlankType.WeldingElement => "Сварное изделие",
        BlankType.Purchased => "Покупная",
        BlankType.Casting => "Литье",
        BlankType.Forging => "Поковка",
        BlankType.CustomBlank => "Не распознано",
        _ => "Не распознано"
    };

    private static string ToTitle(string value) => value.Length == 0 ? value : char.ToUpperInvariant(value[0]) + value[1..].ToLowerInvariant();
    private static string Format(decimal? value) => value is null ? string.Empty : value.Value.ToString("0.####", CultureInfo.InvariantCulture);

    [GeneratedRegex(@"\s+")]
    private static partial Regex SpacesRegex();

    [GeneratedRegex(@"ГОСТ\s*\d{3,6}-\d{4}")]
    private static partial Regex GostRegex();

    [GeneratedRegex(@"(?:D|Ø|Ф)\s*(\d+(?:[\.,]\d+)?)")]
    private static partial Regex DiameterRegex();

    [GeneratedRegex(@"(?:L|ДЛИН[АЫ]?|ДЛ\.)\s*(\d+(?:[\.,]\d+)?)")]
    private static partial Regex LengthRegex();

    [GeneratedRegex(@"(?:S|ТОЛЩИНА)\s*(\d+(?:[\.,]\d+)?)")]
    private static partial Regex ThicknessRegex();

    [GeneratedRegex(@"(?:СТЕНК[АИ]?|СТЕНКА)\s*(\d+(?:[\.,]\d+)?)")]
    private static partial Regex WallRegex();

    [GeneratedRegex(@"(?<n>\d+(?:[\.,]\d+)?)\s*[XХxх×]\s*(?<n>\d+(?:[\.,]\d+)?)(?:\s*[XХxх×]\s*(?<n>\d+(?:[\.,]\d+)?))?")]
    private static partial Regex DimensionsRegex();

    [GeneratedRegex(@"(?<prefix>СТАЛЬ|СТ\.?|НЕРЖ|БРОНЗ[АЫ]?|ЛАТУНЬ)\s*(?<grade>\d{1,3}[А-ЯA-Z]{0,6}\d*[А-ЯA-Z0-9]*|[А-ЯA-Z]{2,}\d{1,3}[А-ЯA-Z0-9]*)", RegexOptions.IgnoreCase)]
    private static partial Regex MaterialRegex();

    [GeneratedRegex(@"\d+(?:[\.,]\d+)?")]
    private static partial Regex NumberRegex();

    [GeneratedRegex(@"\d+\s*[XХxх×]\s*\d+\s*[XХxх×]\s*\d+")]
    private static partial Regex RectPipeRegex();
}
