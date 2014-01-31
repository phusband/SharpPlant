﻿//
//  Copyright © 2013 Parrish Husband (parrish.husband@gmail.com)
//  The MIT License (MIT) - See LICENSE.txt for further details.
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using SharpPlant;

namespace SharpPlant.SharpPlantReview
{
    /// <summary>
    ///     Provides the properties for creating a tag in SmartPlant Review.
    /// </summary>
    public class SprTag
    {
        #region Tag Properties

        /// <summary>
        ///     The parent Application reference.
        /// </summary>
        public SprApplication Application { get; private set; }

        /// <summary>
        ///     Holds the tag bitmask values used for tag placement.
        /// </summary>
        internal int Flags;

        /// <summary>
        ///     Indicates whether a leader will be present.
        /// </summary>
        public bool DisplayLeader
        {
            get
            {
                // Return the bitwise zero check
                return (Flags & SprConstants.SprTagLeader) != 0;
            }
            set
            {
                // Set flag true/false
                if (value)
                    Flags |= SprConstants.SprTagLeader;
                else
                    Flags &= ~SprConstants.SprTagLeader;
            }
        }

        /// <summary>
        ///     Point where the tag is displayed.  If the tag has not been placed, the point coordinates will be 0, 0, 0.
        /// </summary>
        public SprPoint3D OriginPoint
        {
            get
            {
                if (!IsPlaced) return new SprPoint3D(0, 0, 0);
                return new SprPoint3D(Convert.ToDouble(Data["tag_origin_x"]),
                                   Convert.ToDouble(Data["tag_origin_y"]),
                                   Convert.ToDouble(Data["tag_origin_z"]));
            }
            set
            {
                if (!IsPlaced) return;
                Data["tag_origin_x"] = value.East;
                Data["tag_origin_y"] = value.North;
                Data["tag_origin_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Point for the end of the leader.  If the tag has not been placed, the point coordinates will be 0, 0, 0.
        /// </summary>
        public SprPoint3D LeaderPoint
        {
            get
            {
                if (!IsPlaced) return new SprPoint3D(0, 0, 0);
                return new SprPoint3D(Convert.ToDouble(Data["tag_point_x"]),
                                   Convert.ToDouble(Data["tag_point_y"]),
                                   Convert.ToDouble(Data["tag_point_z"]));
            }
            set
            {
                if (!IsPlaced) return;
                Data["tag_point_x"] = value.East;
                Data["tag_point_y"] = value.North;
                Data["tag_point_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Tag text.
        /// </summary>
        public string Text
        {
            get { return (string)Data["tag_text"]; }
            set { Data["tag_text"] = value; }
        }

        /// <summary>
        ///     Size of the tag bubble. If the tag has not been placed, the value is set to zero.
        /// </summary>
        public double Size
        {
            get { return IsPlaced ? Convert.ToDouble(Data["tag_size"]) : 0; }
            set { if (IsPlaced) Data["tag_size"] = value; }
        }

        /// <summary>
        ///     Date the tag was placed.  If the tag has not been placed, the value is N/A.
        /// </summary>
        public string DatePlaced
        {
            get { return IsPlaced ? Data["date_placed"].ToString() : "N/A"; }
            internal set { if (IsPlaced) Data["date_placed"] = value; }
        }

        /// <summary>
        ///     Date the tag was last edited.  If the tag has not been placed, the value is N/A.
        /// </summary>
        public string LastEdited
        {
            get { return IsPlaced ? Data["last_edited"].ToString() : "N/A"; }
            internal set { if (IsPlaced) Data["last_edited"] = value; }
        }

        /// <summary>
        ///     Color of the tag text.
        /// </summary>
        public Color TextColor
        {
            get { return SprUtilities.From0Bgr((int) Data["number_color"]); }
            set { Data["number_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Color of the tag background.
        /// </summary>
        public Color BackgroundColor
        {
            get { return SprUtilities.From0Bgr((int) Data["backgnd_color"]); }
            set { Data["backgnd_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Color of the tag leader line.
        /// </summary>
        public Color LeaderColor
        {
            get { return SprUtilities.From0Bgr((int) Data["leader_color"]); }
            set { Data["leader_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Discipline the tag is set to.
        /// </summary>
        public string Discipline
        {
            get { return (string)Data["discipline"]; }
            set { Data["discipline"] = value; }
        }

        /// <summary>
        ///     Author of the tag.
        /// </summary>
        public string Creator
        {
            get { return (string)Data["creator"]; }
            set { Data["creator"] = value; }
        }

        /// <summary>
        ///     Computer the tag was created on.
        /// </summary>
        public string ComputerName
        {
            get { return (string)Data["computer_name"]; }
            set { Data["computer_name"] = value; }
        }

        /// <summary>
        ///     Status of the tag.
        /// </summary>
        public string Status
        {
            get { return (string)Data["status"]; }
            set { Data["status"] = value; }
        }

        /// <summary>
        ///     Determines if the tag has been placed in SmartPlant Review.
        /// </summary>
        public bool IsPlaced
        { 
            get
            {
                // 42 is the default field count in the tag_data table
                return Data.Count >= 42;
            }
        }

        /// <summary>
        ///     Determines if the tag has an image stored in the MDB.
        /// </summary>
        public bool HasImage
        {
            get
            {
                if (!Data.ContainsKey("tag_image"))
                    return false;
                return Data["tag_image"] != DBNull.Value;
            }
        }

        /// <summary>
        ///     Determines if the tag has labels linked to the MDB2
        /// </summary>
        public bool IsDataLinked
        {
            get { return Linkage.ToString() != "0 0 0 0"; }
        }

        /// <summary>
        ///     The object label linkage owned by the tag.
        /// </summary>
        public SprLinkage Linkage
        {
            get
            {
                var linkString = string.Format("{0} {1} {2} {3}",
                                      Data["linkage_id_0"],
                                      Data["linkage_id_1"],
                                      Data["linkage_id_2"],
                                      Data["linkage_id_3"]);

                return new SprLinkage(linkString);
            }
        }

        /// <summary>
        ///     Tag unique identification number.
        /// </summary>
        public int Id
        {
            get { return Convert.ToInt32(Data["tag_unique_id"]); }
            private set { Data["tag_unique_id"] = value; }
        }

        /// <summary>
        ///     The object the tag is attached to, if any.
        /// </summary>
        public SprObject LinkedObject
        {
            get
            {
                if (!IsDataLinked)
                    return null;

                if (_linkedObject == null)
                    _linkedObject = GetLinkedObject();

                return _linkedObject;
                //return GetLinkedObject();
            }
        }
        private SprObject _linkedObject;

        /// <summary>
        ///     The full information profile of the current tag.  Controls all the tag properties.
        /// </summary>
        public Dictionary<string, object> Data { get; set; }

        #endregion

        // Tag constructor
        public SprTag()
        {
            // Link the parent application
            Application = SprApplication.ActiveApplication;

            // Create a new data dictionary from the template
            Data = SprUtilities.TagTemplate();

            // Set the tag to the next available tag number
            Id = Application.NextTag;

            // Set the leader display by default
            DisplayLeader = true;

            // Add the default data values
            BackgroundColor = SprUtilities.From0Bgr(12632319);
            LeaderColor = SprUtilities.From0Bgr(12632319);

            // Set the tag creator
            Creator = Environment.GetEnvironmentVariable("USERNAME");

            // Set the computer name
            ComputerName = Environment.GetEnvironmentVariable("COMPUTERNAME");

            // Set the default status
            Status = "Open";
        }

        private SprObject GetLinkedObject()
        {
            var searchString = string.Format("FIND linkages = {0}",Linkage.ToString());
            var objIds = Application.ObjectDataSearch(searchString);
            if (objIds.Count > 1)
            {
                // Use the tag origin point to filter down to a single object
                var volumeFormat = "{0} N, {1} E, {2} El";
                var volumeStart = string.Format(volumeFormat, OriginPoint.North, OriginPoint.East, OriginPoint.Elevation);
                var volumeEnd = string.Format(volumeFormat, OriginPoint.North, OriginPoint.East, OriginPoint.Elevation);

                searchString = string.Format("{0}\nKEEP ONLY Volume Overlap {1} to {2}",
                                            searchString, volumeStart, volumeEnd);

                objIds = Application.ObjectDataSearch(searchString);
            }

            if (objIds.Count == 0)
                return null;
            else if (objIds.Count > 1)
                throw new SprException("Multiple objects found for linkage " + Linkage.ToString());
            return Application.GetObjectData(objIds[0]);         
        }
    }
}