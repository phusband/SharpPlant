﻿//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Controls tag visibility in the main view.
    /// </summary>
    public enum SprTagVisibility
    {
        // Each value corresponds to the ASCII character code
        /// <summary>
        ///     No tags displayed.
        /// </summary>
        None = 78, // N
        /// <summary>
        ///     Only the active tag displayed.
        /// </summary>
        ActiveOnly = 86, // V
        /// <summary>
        ///     All tags displayed.
        /// </summary>
        All = 76 // L
    }
}
