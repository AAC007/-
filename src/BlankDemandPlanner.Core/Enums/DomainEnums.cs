namespace BlankDemandPlanner.Core.Enums;

public enum BlankType
{
    RoundBar,
    SquareBar,
    HexBar,
    Sheet,
    Plate,
    PipeRound,
    PipeRectangular,
    Angle,
    Channel,
    IBeam,
    BronzeBar,
    BronzeSheet,
    CustomBlank,
    Unknown,
    WeldingElement,
    Purchased,
    Casting,
    Forging
}

public enum I012Status
{
    Allowed,
    SizeNotAllowed,
    MaterialNotAllowedForSize,
    TypeNotFound,
    Unrecognized
}

public enum MeasurementUnit
{
    Piece,
    Meter,
    Kilogram
}

public enum CalculationStatus
{
    Ok,
    MissingPart,
    MissingBlankMapping,
    MultipleActiveMappings,
    UnitMismatch
}

public enum ImportRunStatus
{
    Started,
    Completed,
    CompletedWithErrors,
    Failed,
    Cancelled
}

public enum SortDirection
{
    Ascending,
    Descending
}
