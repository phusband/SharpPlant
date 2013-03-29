using System;

namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Contains information about a particular measurement in the active Measurement Collection.
    /// </summary>
    public class Measurement
    {
        #region Measurement Properties

        /// <summary>
        ///     The active COM reference to the DrMeasurement class
        /// </summary>
        internal dynamic DrMeasurement;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public Application Application { get; private set; }

        /// <summary>
        ///     Distance between all previous points and this point.
        /// </summary>
        public double CumulativeDistance
        {
            get
            {
                if (IsActive) return DrMeasurement.CumulativeDistance;
                return -1;
            }
        }

        /// <summary>
        ///     Distance between the previous point and this point.
        /// </summary>
        public double Distance
        {
            get
            {
                if (IsActive) return DrMeasurement.Distance;
                return -1;
            }
        }

        /// <summary>
        ///     The index of this measurement in the current measurement collection.
        /// </summary>
        public int Index
        {
            get
            {
                if (IsActive) return DrMeasurement.Index;
                return -1;
            }
        }

        /// <summary>
        ///     Primary point coordinate of the measurement.
        /// </summary>
        public Point3D ActivePoint
        {
            get
            {
                if (IsActive)
                    return new Point3D(DrMeasurement.Point.East,
                                       DrMeasurement.Point.North,
                                       DrMeasurement.Point.Elevation);
                return null;
            }
            set
            {
                if (!IsActive) return;
                DrMeasurement.Point.East = value.East;
                DrMeasurement.Point.North = value.North;
                DrMeasurement.Point.Elevation = value.Elevation;
            }
        }

        /// <summary>
        ///     Prevous measurement point (if any).
        /// </summary>
        public Point3D LastPoint
        {
            get
            {
                if (IsActive)
                    return new Point3D(DrMeasurement.LastPoint.East,
                                       DrMeasurement.LastPoint.North,
                                       DrMeasurement.LastPoint.Elevation);
                return null;
            }
            set
            {
                if (!IsActive) return;
                DrMeasurement.LastPoint.East = value.East;
                DrMeasurement.LastPoint.North = value.North;
                DrMeasurement.LastPoint.Elevation = value.Elevation;
            }
        }

        /// <summary>
        ///     The Point coordinate of the text label location.
        /// </summary>
        public Point3D TextPoint
        {
            get
            {
                if (IsActive)
                    return new Point3D(DrMeasurement.TextPoint.East,
                                       DrMeasurement.TextPoint.North,
                                       DrMeasurement.TextPoint.Elevation);
                return null;
            }
            set
            {
                if (!IsActive) return;
                DrMeasurement.TextPoint.East = value.East;
                DrMeasurement.TextPoint.North = value.North;
                DrMeasurement.TextPoint.Elevation = value.Elevation;
            }
        }

        /// <summary>
        ///     Measurement type setting.
        /// </summary>
        public MeasurementType Type
        {
            get { return IsActive ? (MeasurementType) DrMeasurement.Type : MeasurementType.Null; }
            set { if (IsActive) DrMeasurement.Type = (int) value; }
        }

        /// <summary>
        ///     True if the measurement range is valid, false if not.
        /// </summary>
        public bool IsRangeValid
        {
            get
            {
                if (IsActive) return (bool) DrMeasurement.RangeValid;
                return false;
            }
        }

        // Determines whether a reference to the COM object is established
        private bool IsActive
        {
            get { return (DrMeasurement != null); }
        }

        #endregion

        // Measurement initializer
        public Measurement()
        {
            // Link the parent application
            Application = SmartPlantReview.ActiveApplication;

            // Get a new DrPointDbl object
            DrMeasurement = Activator.CreateInstance(ImportedTypes.DrMeasurement);
        }
    }
}