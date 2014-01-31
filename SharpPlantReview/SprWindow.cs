//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the structure containing information about a SmartPlant Review window.
    /// </summary>
    public class SprWindow
    {
        #region SPRWindow Properties

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     Height of the working area of the window.
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        ///     Left(x) position of the top-left corner of the window.
        /// </summary>
        public int Left { get; set; }

        /// <summary>
        ///     Top (y) position of the top-left corner of the window.
        /// </summary>
        public int Top { get; set; }

        /// <summary>
        ///     Width of the working area of the window.
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        ///     HWND of the window.
        /// </summary>
        public int WindowHandle { get; internal set; }

        /// <summary>
        ///     Index of the window.
        /// </summary>
        public int Index { get; internal set; }

        #endregion


        // SPRWindow initializer
        internal SprWindow()
        {
            // Link the parent application
            Application = SprApplication.ActiveApplication;
        }
    }
}