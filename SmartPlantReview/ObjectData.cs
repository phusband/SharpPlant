using System;
using System.Collections.Generic;

namespace SharpPlant.SmartPlantReview
{
    /// <summary>
    ///     Contains information about a particular SmartPlant Review object.
    /// </summary>
    public class ObjectData
    {
        #region ObjectData Properties

        /// <summary>
        ///     The active COM reference to the DrObjectDataDbl class
        /// </summary>
        internal dynamic DrObjectDataDbl;

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public Application Application { get; private set; }

        /// <summary>
        ///     DataObject color index.
        /// </summary>
        public int Color
        {
            get
            {
                if (IsActive) return DrObjectDataDbl.Color;
                return -1;
            }
        }

        /// <summary>
        ///     Number of display sets the ObjectData object belongs to.
        /// </summary>
        public int DisplaySetCount
        {
            get
            {
                if (IsActive) return DrObjectDataDbl.DisplaySetCount;
                return -1;
            }
        }

        /// <summary>
        ///     Display set id of highest priority display set, 0 if no display set.
        /// </summary>
        public int DisplaySetId
        {
            get
            {
                if (IsActive) return DrObjectDataDbl.DisplaySetID;
                return -1;
            }
        }

        /// <summary>
        ///     Display set name of highest priority.
        /// </summary>
        public string DisplaySetName
        {
            get
            {
                if (IsActive)
                    return DrObjectDataDbl.DisplaySetName == "0" ? string.Empty : DrObjectDataDbl.DisplaySetName;
                return string.Empty;
            }
        }

        /// <summary>
        ///     Name of the file that the ObjectData object is in.
        /// </summary>
        public string FileName
        {
            get
            {
                return IsActive ? DrObjectDataDbl.FileName : string.Empty;
            }
        }

        /// <summary>
        ///     Name of the ObjectData label file.
        /// </summary>
        public string LabelFileName
        {
            get
            {
                if (IsActive)
                    return DrObjectDataDbl.LabelFileName == "0" ? string.Empty : DrObjectDataDbl.LabelFileName;
                return string.Empty;
            }
        }

        /// <summary>
        ///     DataObject level number.
        /// </summary>
        public int Level
        {
            get
            {
                if (IsActive) return DrObjectDataDbl.Level;
                return -1;
            }
        }

        /// <summary>
        ///     DataObject material name.
        /// </summary>
        public string MaterialName
        {
            get
            {
                if (IsActive)
                    return DrObjectDataDbl.MaterialName == "0" ? string.Empty : DrObjectDataDbl.MaterialName;
                return string.Empty;
            }
        }

        /// <summary>
        ///     DataObject palette name.
        /// </summary>
        public string PaletteName
        {
            get
            {
                if (IsActive)
                    return DrObjectDataDbl.PaletteName == "0" ? string.Empty : DrObjectDataDbl.PaletteName;
                return string.Empty;
            }
        }

        /// <summary>
        ///     The Object ID of the selected ObjectData.
        /// </summary>
        public int ObjectId { get; internal set; }

        /// <summary>
        ///     The 3D point where the ObjectData was selected.
        /// </summary>
        public Point3D SelectedPoint { get; internal set; }

        /// <summary>
        ///     A collection of key/value pairs containing object label entries
        /// </summary>
        public Dictionary<string, string> LabelData { get; private set; }

        // Determines whether a reference to the COM object is established
        private bool IsActive
        {
            get { return (DrObjectDataDbl != null); }
        }

        #endregion

        // DataObject initializer
        internal ObjectData()
        {
            // Link the parent application
            Application = SmartPlantReview.ActiveApplication;

            // Get a new DrObjectDataDbl object
            DrObjectDataDbl = Activator.CreateInstance(ImportedTypes.DrObjectDataDbl);

            // Create the label dictionary
            LabelData = new Dictionary<string, string>();
        }
    }
}