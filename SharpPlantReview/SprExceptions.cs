//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;

namespace SharpPlant.SharpPlantReview
{
    internal static class SprExceptions
    {
        internal static Exception SprNotConnected = new Exception("A SmartPlant Review connection is not established.");
        internal static Exception SprOutOfMemory = new Exception("The SmartPlant Review application ran out of memory.");
        internal static Exception SprObjectCreateFail = new Exception("The SmartPlant API object failed to be instanciated.");
        internal static Exception SprTagNotPlaced = new Exception("The provided tag has not been placed in SmartPlant Review.");
        internal static Exception SprNullPoint = new Exception("One or more required points are null.");
        internal static Exception SprTagNotFound = new Exception("The desired tag does not exist in the MDB database.");
        internal static Exception SprAnnotationNotFound = new Exception("The desired annotation does not exist in the MDB database.");
        internal static Exception SprUnsupportedMethod = new Exception("The active version of SmartPlant Review does not suppoort this method.");
    }
}