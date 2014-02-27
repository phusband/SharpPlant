//
//  Copyright © 2014 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Contains information about a particular SmartPlant Review object.
    /// </summary>
    public class SprObject : IEquatable<SprObject>
    {
        #region Properties

        /// <summary>
        ///     The active COM reference to the DrObjectDataDbl class
        /// </summary>
        internal dynamic DrObjectDataDbl;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     SprObject color index.
        /// </summary>
        public int Color
        {
            get { return IsActive ? DrObjectDataDbl.Color : -1; }
        }

        /// <summary>
        ///     Number of display sets the SprObject belongs to.
        /// </summary>
        public int DisplaySetCount
        {
            get { return IsActive ? DrObjectDataDbl.DisplaySetCount : -1; }
        }

        /// <summary>
        ///     Display set id of highest priority display set, 0 if no display set.
        /// </summary>
        public int DisplaySetId
        {
            get { return IsActive ? DrObjectDataDbl.DisplaySetID : -1; }
        }

        /// <summary>
        ///     Display set name of highest priority.
        /// </summary>
        public string DisplaySetName
        {
            get
            {
                return IsActive
                    ? (DrObjectDataDbl.DisplaySetName == "0" ? string.Empty
                    : DrObjectDataDbl.DisplaySetName)
                    : string.Empty;
            }
        }

        /// <summary>
        ///     Name of the file that the SprObject is in.
        /// </summary>
        public string FileName
        {
            get
            {
                return IsActive ? DrObjectDataDbl.FileName : string.Empty;
            }
        }

        /// <summary>
        ///     Name of the SprObject label file.
        /// </summary>
        public string LabelFileName
        {
            get
            {
                return IsActive
                    ? (DrObjectDataDbl.LabelFileName == "0" ? string.Empty
                    : DrObjectDataDbl.LabelFileName)
                    : string.Empty;
            }
        }

        /// <summary>
        ///     The object label linkage.
        /// </summary>
        public SprLinkage Linkage
        {
            get
            {
                return IsActive ? new SprLinkage(DrObjectDataDbl.LabelKey) : null;
            }
        }

        /// <summary>
        ///     The level number of the SprObject.
        /// </summary>
        public int Level
        {
            get { return IsActive ? DrObjectDataDbl.Level : -1; }
        }

        /// <summary>
        ///     The material name of the SprObject.
        /// </summary>
        public string MaterialName
        {
            get
            {
                return IsActive
                    ? (DrObjectDataDbl.MaterialName == "0" ? string.Empty
                    : DrObjectDataDbl.MaterialName)
                    : string.Empty;
            }
        }

        /// <summary>
        ///     The material palette file of the SprObject.
        /// </summary>
        public string PaletteName
        {
            get
            {
                return IsActive
                    ? (DrObjectDataDbl.PaletteName == "0" ? string.Empty
                    : DrObjectDataDbl.PaletteName)
                    : string.Empty;
            }
        }

        /// <summary>
        ///     The Object Id of the SprObject.
        /// </summary>
        public int Id { get; internal set; }

        /// <summary>
        ///     A collection of key/value pairs containing object label entries
        /// </summary>
        public Dictionary<string, string> Labels { get; private set; }

        // Determines whether a reference to the COM object is established
        private bool IsActive
        {
            get { return (DrObjectDataDbl != null); }
        }

        #endregion

        // DataObject initializer
        internal SprObject()
        {
            // Link the parent application
            Application = SprApplication.ActiveApplication;

            // Get a new DrObjectDataDbl object
            DrObjectDataDbl = Activator.CreateInstance(SprImportedTypes.DrObjectDataDbl);

            // Create the label dictionary
            Labels = new Dictionary<string, string>();
        }

        #region IEquatable

        public bool Equals(SprObject other)
        {
            if (ReferenceEquals(other, null))
                return false;

            if (ReferenceEquals(this, null))
                return true;

            return other.Linkage.Equals(Linkage) &&
                   other.LabelFileName.Equals(LabelFileName) &&
                   other.FileName.Equals(FileName);
        }

        public override bool Equals(Object obj)
        {
            return Equals(obj as SprObject);
        }

        public override int GetHashCode()
        {
            return Equals(null)
                ? 0
                : new {Linkage, LabelFileName, FileName}.GetHashCode();
        }

        public static bool operator ==(SprObject left, SprObject right)
        {
            return ReferenceEquals(left, null)
                ? ReferenceEquals(right, null)
                : left.Equals(right);
        }

        public static bool operator !=(SprObject left, SprObject right)
        {
            return !(left == right);
        }

        #endregion
    }
}