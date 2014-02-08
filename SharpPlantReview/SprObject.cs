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
    public class SprObject
    {
        #region SprObject Properties

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
            get
            {
                if (IsActive) return DrObjectDataDbl.Color;
                return -1;
            }
        }

        /// <summary>
        ///     Number of display sets the SprObject belongs to.
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
                if (IsActive)
                    return DrObjectDataDbl.LabelFileName == "0" ? string.Empty : DrObjectDataDbl.LabelFileName;
                return string.Empty;
            }
        }

        /// <summary>
        ///     The object label linkage.
        /// </summary>
        public SprLinkage Linkage
        {
            get
            {
                if (IsActive)
                    return new SprLinkage(DrObjectDataDbl.LabelKey);
                return null;
            }
        }

        /// <summary>
        ///     The level number of the SprObject.
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
        ///     The material name of the SprObject.
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
        ///     The material palette file of the SprObject.
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
    }
}