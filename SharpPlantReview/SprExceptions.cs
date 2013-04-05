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
        internal static Exception SprInvalidFileName = new Exception("An invalid filename was provided.");
        internal static Exception SprInvalidFileNumber = new Exception("An invalid file number exception occurred.");
        internal static Exception SprInvalidParameter = new Exception("An invalid parameter was provided.");
        internal static Exception SprInternalError = new Exception("An internal error occured inside SmartPlant Review.");
        internal static Exception SprInvalidObjectId = new Exception("The provided ObjectId was invalid.");
        internal static Exception SprOutOfMemory = new Exception("The SmartPlant Review application ran out of memory.");
        internal static Exception SprObjectCreateFail = new Exception("The SmartPlant API object failed to be instanciated.");
        internal static Exception SprInvalidTag = new Exception("An invalid tag number was provided.");
        internal static Exception SprTagExists = new Exception("The tag number already exists.");
        internal static Exception SprInvalidView = new Exception("The provided view was invalid.");
        internal static Exception SprInvalidGlobal = new Exception("The global option specified was invalid");
        internal static Exception SprInvalidAnnoType = new Exception("The annotation type was invalid.");
        internal static Exception SprDirectoryWriteFailure = new Exception("The write attempt to the specified directory failed.");
        internal static Exception SprFileWriteFailure = new Exception("The write attempt to the specified file failed.");
        internal static Exception SprFileExists = new Exception("The specified file already exists.");
        internal static Exception SprTagNotPlaced = new Exception("The provided tag has not been placed in SmartPlant Review.");
        internal static Exception SprNullPoint = new Exception("One or more required points are null.");
        internal static Exception SprTagNotFound = new Exception("The desired tag does not exist in the Mdb database.");
    }
}