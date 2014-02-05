//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides properties for controlling Display Sets in SmartPlant Review.
    /// </summary>
    public class SprDisplaySet
    {
        /// <summary>
        ///     The active COM reference to the DrDisplaySetDbl class
        /// </summary>
        internal dynamic DrDisplaySetDbl;

        //public double 


        public SprDisplaySet()
        {
            DrDisplaySetDbl = Activator.CreateInstance(SprImportedTypes.DrDisplaySet);
        }
    }
}
