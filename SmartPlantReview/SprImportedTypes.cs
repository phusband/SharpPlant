using System;

namespace SharpPlant.SmartPlantReview
{
    internal static class SprImportedTypes
    {
        internal static Type DrApi = Type.GetTypeFromProgID("DrApi.DrApi.1");
        internal static Type DrAnnotationDbl = Type.GetTypeFromProgID("DrAnnotationDbl.DrAnnotationDbl.1");
        internal static Type DrDisplaySet = Type.GetTypeFromProgID("DrDisplaySet.DrDisplaySet.1");
        internal static Type DrKey = Type.GetTypeFromProgID("DrKey.DrKey.1");
        internal static Type DrLongArray = Type.GetTypeFromProgID("DrLongArray.DrLongArray.1");
        internal static Type DrMask = Type.GetTypeFromProgID("DrMask.DrMask.1");
        internal static Type DrMeasurement = Type.GetTypeFromProgID("DrMeasurement.DrMeasurement.1");

        internal static Type DrMeasurementCollection =
            Type.GetTypeFromProgID("DrMeasurementCollection.DrMeasurementCollection.1");

        internal static Type DrObjectDataDbl = Type.GetTypeFromProgID("DrObjectDataDbl.DrObjectDataDbl.1");
        internal static Type DrPointDbl = Type.GetTypeFromProgID("DrPointDbl.DrPointDbl.1");
        internal static Type DrSnapShot = Type.GetTypeFromProgID("DrSnapShot.DrSnapShot.1");
        internal static Type DrStringArray = Type.GetTypeFromProgID("DrStringArray.DrStringArray.1");
        internal static Type DrTransform = Type.GetTypeFromProgID("DrTransform.DrTransform.1");
        internal static Type DrViewDbl = Type.GetTypeFromProgID("DrViewDbl.DrViewDbl.1");
        internal static Type DrVolumeAnnotation = Type.GetTypeFromProgID("DrVolumeAnnotation.DrVolumeAnnotation.1");
        internal static Type DrWindow = Type.GetTypeFromProgID("DrWindow.DrWindow.1");
    }
}