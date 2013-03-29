using System;

namespace SharpPlant.SmartPlantReview
{
    internal static class Exceptions
    {
        internal static Exception ApiNotConnected = new Exception("A SmartPlant Review connection is not established");
    }
}