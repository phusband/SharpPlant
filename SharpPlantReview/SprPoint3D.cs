//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the structure for a 3D point in SmartPlant Review.
    /// </summary>
    public class SprPoint3D : IEquatable<SprPoint3D>
    {
        #region Properties

        /// <summary>
        ///     Active COM reference to the DrPointDbl class.
        /// </summary>
        internal dynamic DrPointDbl;

        /// <summary>
        ///     Easting coordinate
        /// </summary>
        public double East
        {
            get { return IsActive ? DrPointDbl.East : -1; }
            set { if (IsActive) DrPointDbl.East = value; }
        }

        /// <summary>
        ///     Northing coordinate
        /// </summary>
        public double North
        {
            get { return IsActive ? DrPointDbl.North : -1; }
            set { if (IsActive) DrPointDbl.North = value; }
        }

        /// <summary>
        ///     Elevation coordinate
        /// </summary>
        public double Elevation
        {
            get { return IsActive ? DrPointDbl.Elevation : -1; }
            set { if (IsActive) DrPointDbl.Elevation = value; }
        }

        // Determines if a reference to the COM object is established
        private bool IsActive
        {
            get { return DrPointDbl != null; }
        }

        #endregion

        #region Constructors

        internal SprPoint3D(dynamic drPointDbl)
        {
            DrPointDbl = drPointDbl;
        }
        public SprPoint3D()
        {
            DrPointDbl = System.Activator.CreateInstance(SprImportedTypes.DrPointDbl);
            if (DrPointDbl == null) throw SprExceptions.SprObjectCreateFail;
        }
        public SprPoint3D(double east, double north, double elevation)
        {
            DrPointDbl = System.Activator.CreateInstance(SprImportedTypes.DrPointDbl);
            if (DrPointDbl == null) throw SprExceptions.SprObjectCreateFail;

            East = east;
            North = north;
            Elevation = elevation;
        }

        #endregion

        public bool Equals(SprPoint3D other)
        {
            if (ReferenceEquals(other, null))
                return false;

            if (ReferenceEquals(this, null))
                return true;

            return other.North.Equals(North) &&
                   other.East.Equals(East) &&
                   other.Elevation.Equals(Elevation);
        }

        public override int GetHashCode()
        {
            return Equals(null)
                ? 0
                : new {East, North, Elevation}.GetHashCode();
        }

        public override bool Equals(Object obj)
        {
            return Equals(obj as SprPoint3D);
        }

        public static bool operator ==(SprPoint3D left, SprPoint3D right)
        {
            return ReferenceEquals(left, null) 
                ? ReferenceEquals(right, null) 
                : left.Equals(right);
        }

        public static bool operator !=(SprPoint3D left, SprPoint3D right)
        {
            return !(left == right);
        }
    }
}