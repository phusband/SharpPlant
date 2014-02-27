//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Runtime.Serialization;

namespace SharpPlant.SharpPlantReview
{
    [Serializable]
    public class SprException : Exception
    {
        public SprException(string message, params object[] args) : base(string.Format(message, args)) 
        { }

        // Ensure Exception is Serializable
        protected SprException(SerializationInfo info, StreamingContext ctxt) : base(info, ctxt)
        { }
    }

    internal static class SprExceptions
    {
        internal static SprException SprNotConnected = new SprException("A SmartPlant Review connection is not established.");
        internal static SprException SprObjectCreateFail = new SprException("The SmartPlant API object failed to be instanciated.");
        internal static SprException SprTagNotPlaced = new SprException("The provided tag has not been placed in SmartPlant Review.");
        internal static SprException SprNullPoint = new SprException("One or more required points are null.");
        internal static SprException SprTagNotFound = new SprException("The desired tag does not exist in the MDB database.");
        internal static SprException SprAnnotationNotFound = new SprException("The desired annotation does not exist in the MDB database.");
        internal static SprException SprVersionIncompatibility = new SprException("The operation is not supported in the current version of SmartPlant Review.");
        //internal static SprException SprMdbAccess = new SprException();
    }
}