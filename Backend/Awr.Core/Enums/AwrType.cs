namespace Awr.Core.Enums
{
    // The fixed set of AWR document types that can be requested
    public enum AwrType
    {
        Others = 0,     // Replaces 'Unknown' for valid, non-listed types
        FPS = 1,        // Final Product Specification
        IMS = 2,        // In-Process Material Specification
        MICRO = 3,      // Microbiology
        PM = 4,         // Packaging Material
        RM = 5,         // Raw Material
        STABILITY = 6,  // Stability Testing
        WATER = 7       // Water Testing
    }
}