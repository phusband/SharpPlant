namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Contains application constants used by SmartPlant Review.
    /// </summary>
    public static class Constants
    {
        // General
        public const int TEXT_WIN_CLR = 1;

        // Snapshot settings
        public const int FILE_OVERWRITE_OK = 1;
        public const int SNAP_FULL_SCREEN = 2;
        public const int SNAP_ROTATE_90 = 4;
        public const int SNAP_ASPECT_ON = 8;
        public const int SNAP_RESERVED1 = 16;
        public const int SNAP_FORCE_BMP = 32;
        public const int SNAP_FORCE_RGB = 64;
        public const int SNAP_ENTIRE_MONITOR = 128;

        // Annotation settings
        public const int ANNO_LEADER = 1;
        public const int ANNO_ARROW = 2;
        public const int ANNO_BACKGROUND = 4;
        public const int ANNO_DEFAULT_COLORS = 8;
        public const int ANNO_PERSIST = 16;

        // Tag settings
        public const int TAG_LEADER = 1;
        public const int TAG_LABEL_KEY = 2;
        public const int TAG_EDIT_OK = 4;
        public const int TAG_RESERVED1 = 8;
        public const int TAG_RESERVED2 = 16;

        // Global options
        public const int G_ANNO_DISPLAYED = 8;
        public const int G_ANNO_TEXT_DISPLAYED = 9;
        public const int G_API_FILE_INFO_MODE = 47;

        // Views
        public const int MAIN_VIEW = 0;
        public const int PLAN_VIEW = 1;
        public const int ELEV_VIEW = 2;

        // Windows
        public const int MAIN_WIN = 0;
        public const int PLAN_WIN = 1;
        public const int ELEV_WIN = 2;
        public const int TEXT_WIN = 4;
        public const int SMARTPLANT_REVIEW_WIN = 5;
    }
}