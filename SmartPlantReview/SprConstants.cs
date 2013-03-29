namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Contains application constants used by SmartPlant Review.
    /// </summary>
    public static class SprConstants
    {
        // General
        public const int SprClearTextWindow = 1;

        // Snapshot settings
        public const int SprSnapOverwrite = 1;
        public const int SprSnapFullscreen = 2;
        public const int SprSnapRotate90 = 4;
        public const int SprSnapAspectOn = 8;
        public const int SprSnapForceBmp = 32;
        public const int SprSnapForceRgb = 64;
        public const int SprSnapEntireMonitor = 128;

        // Annotation settings
        public const int SprAnnoLeader = 1;
        public const int SprAnnoArrow = 2;
        public const int SprAnnoBackground = 4;
        public const int SprAnnoDefaultColor = 8;
        public const int SprAnnoPersist = 16;

        // Tag settings
        public const int SprTagLeader = 1;
        public const int SprTagLabel = 2;
        public const int SprTagEdit = 4;

        // Global options
        public const int SprGlobalAnnoDisplay = 8;
        public const int SprGlobalTextDisplay = 9;
        public const int SprGlobalFileInfoMode = 47;

        // Views
        public const int SprMainView = 0;
        public const int SprPlanView = 1;
        public const int SprElevationView = 2;

        // Windows
        public const int SprMainWindow = 0;
        public const int SprPlanWindow = 1;
        public const int SprElevationWindow = 2;
        public const int SprTextWindow = 4;
        public const int SprApplicationWindow = 5;
    }
}