using System;

namespace Awr.Worker.Configuration
{
    public static class DocumentLayout
    {
        // 1. PAGE MARGINS (0.8 cm)
        public const float PageMarginPt = 22.68f;
        public const float HeaderDistPt = 22.68f;
        public const float FooterDistPt = 22.68f;

        // 2. IMAGE RESIZING BOUNDING BOX
        // Target Width: 18.0 cm.
        public const float TargetWidthPt = 510.24f;

        // Target Height: 25.0 cm (Lowered to ensure footer is 100% safe)
        // Calculation: (25.0 / 2.54) * 72 = 708.66 pt
        public const float TargetHeightPt = 708.66f;
    }
}