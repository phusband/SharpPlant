//
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
                try { return Data["tag_image"] != DBNull.Value; }
                catch (KeyNotFoundException){ return false; } 
            }
        }

        /// <summary>
        ///     Determines if the tag has labels linked to the MDB2
        /// </summary>
        public bool IsDataLinked
        {
            get
            {
                return !(Convert.ToInt32(Data["linkage_id_0"]) == 0 &&
                         Convert.ToInt32(Data["linkage_id_1"]) == 0 &&
                         Convert.ToInt32(Data["linkage_id_2"]) == 0 &&
                         Convert.ToInt32(Data["linkage_id_3"]) == 0);
                
                //return Labels != null;
            }
        }

        /// <summary>
        ///     The Key/Value label collection associated with the tag.
        /// </summary>
        public Dictionary<string,string> Labels
        {
            get
            {
                if (!IsDataLinked)
                    return null;
                return GetLabels();
            }
        }

        /// <summary>
        ///     The model number linked to the tagged object.
        /// </summary>
        public string ModelNumber
        {
            get
            {
                if (_modelNumber == null)
                {
                    _modelNumber = GetModelNumber();
                }

                return _modelNumber;
            }
        }
        private string _modelNumber;

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
        //public SprObjectData AssociatedObject
        //{ 
        //    get
        //    {
        //        try { return Application.GetObjectData(Convert.ToInt32(Data["object_id"])); }
        //        catch (KeyNotFoundException){ return null; }
        //        catch (InvalidCastException) { return null; } 
        //    }
        //    internal set
        //    {
        //        try { Data["object_id"] = value.ObjectId; }
        //        catch (KeyNotFoundException) { } 
        //    }
        //}

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

            // Set the backing properties to null by default
            _modelNumber = null;
        }

        // Query the MDB2 database for the labels
        private Dictionary<string,string> GetLabels()
        {
            // Create the return hashtable
            var returnData = new Dictionary<string, string>();

            // Build the linkage address
            var linkID = string.Format("{0} {1} {2} {3}", Data["linkage_id_0"], Data["linkage_id_1"],
                                                          Data["linkage_id_2"], Data["linkage_id_3"]);
            // Create the connection to the MDB2 database
            using (var connection = DbMethods.GetConnection(Application.MdbPath + 2, DbMethods.ConnectionType.Ace))
            {
                // Open the MDB2 connection
                DbMethods.CheckOpenConnection(connection);

                // Create a command
                var command = connection.CreateCommand();

                // Build the commandstring
                command.CommandText = string.Format("SELECT linkage_index FROM linkage WHERE DMRSLinkage = '{0}'", linkID);
                
                // Get the linkage address
                object linkageIndex = null;
                try
                {
                    linkageIndex = command.ExecuteScalar();
                }
                catch (System.Data.OleDb.OleDbException)
                {
                }

                // Return if the linkage was not found
                if (linkageIndex == null)
                    return null;

                // Build the string to gather label information
                command.CommandText = "SELECT label_names.label_name, label_values.label_value " +
                                      "FROM (labels INNER JOIN label_names ON labels.label_name_index = label_names.label_name_index) " +
                                      "INNER JOIN label_values ON labels.label_value_index = label_values.label_value_index " +
                                      "WHERE(((labels.linkage_index) = " + linkageIndex + ")) " +
                                      "ORDER BY labels.label_line_number";

                // Create a datareader to walk through the label data
                using (var reader = command.ExecuteReader())
                {
                    // Iterate through the end of the data
                    while (reader.Read())
                    {
                        // Set the key/value
                        var key = reader.GetString(0);
                        var value = reader.GetString(1);

                        // Add non-duplicate keys to the hashtable
                        if (!returnData.ContainsKey(key))
                        {
                            returnData.Add(key, value);
                        }
                    }
                }
            }

            // Return the hashtable
            if (returnData.Count > 0)
                return returnData;
            return null;

        }

        // Query the MDB2 database for the model number
        private string GetModelNumber()
        {
            // Build the linkage address
            var linkID = string.Format("{0} {1} {2} {3}", Data["linkage_id_0"], Data["linkage_id_1"],
                                                          Data["linkage_id_2"], Data["linkage_id_3"]);

            // Create a connection to the MDB2
            using (var connection = DbMethods.GetConnection(Application.MdbPath + 2, DbMethods.ConnectionType.Ace))
            {
                // Open the MDB2 connection
                DbMethods.CheckOpenConnection(connection);

                // Build the command
                var command = connection.CreateCommand();
                command.CommandText = string.Format("SELECT file_name FROM linkage WHERE DMRSLinkage = '{0}'", linkID);
                
                // Get the model number
                var modelNumber = (string)command.ExecuteScalar();
                
                // Return the model number
                return modelNumber == null ? "Not Found" : modelNumber.ToString();
            }
        }
    }
}