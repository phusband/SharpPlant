using System;
using System.Collections.Generic;
using System.Drawing;

namespace SharpPlant.SmartPlantReview
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
        ///     Point where the tag is displayed.
        /// </summary>
        public SprPoint3D OriginPoint
        {
            get
            {
                return new SprPoint3D(Convert.ToDouble(TagData["tag_origin_x"]),
                                   Convert.ToDouble(TagData["tag_origin_y"]),
                                   Convert.ToDouble(TagData["tag_origin_z"]));
            }
            set
            {
                TagData["tag_origin_x"] = value.East;
                TagData["tag_origin_y"] = value.North;
                TagData["tag_origin_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Point for the end of the leader.
        /// </summary>
        public SprPoint3D LeaderPoint
        {
            get
            {
                return new SprPoint3D(Convert.ToDouble(TagData["tag_point_x"]),
                                   Convert.ToDouble(TagData["tag_point_y"]),
                                   Convert.ToDouble(TagData["tag_point_z"]));
            }
            set
            {
                TagData["tag_point_x"] = value.East;
                TagData["tag_point_y"] = value.North;
                TagData["tag_point_z"] = value.Elevation;
            }
        }

        /// <summary>
        ///     Tag text.
        /// </summary>
        public string Text
        {
            get { return TagData["tag_text"].ToString(); }
            set { TagData["tag_text"] = value; }
        }

        /// <summary>
        ///     Size of the tag bubble.
        /// </summary>
        public double Size
        {
            get { return Convert.ToDouble(TagData["tag_size"]); }
            set { TagData["tag_size"] = value; }
        }

        /// <summary>
        ///     Date the tag was placed.
        /// </summary>
        public string DatePlaced
        {
            get { return TagData["date_placed"].ToString(); }
        }

        /// <summary>
        ///     Date the tag was last edited.
        /// </summary>
        public string LastEdited
        {
            get { return TagData["last_edited"].ToString(); }
        }

        /// <summary>
        ///     Color of the tag text.
        /// </summary>
        public Color TextColor
        {
            get { return SprUtilities.From0Bgr((int) TagData["number_color"]); }
            set { TagData["number_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Color of the tag background.
        /// </summary>
        public Color BackgroundColor
        {
            get { return SprUtilities.From0Bgr((int) TagData["backgnd_color"]); }
            set { TagData["backgnd_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Color of the tag leader line.
        /// </summary>
        public Color LeaderColor
        {
            get { return SprUtilities.From0Bgr((int) TagData["leader_color"]); }
            set { TagData["leader_color"] = SprUtilities.Get0Bgr(value); }
        }

        /// <summary>
        ///     Discipline the tag is set to.
        /// </summary>
        public string Discipline
        {
            get { return TagData["discipline"].ToString(); }
            set { TagData["discipline"] = value; }
        }

        /// <summary>
        ///     Tag unique identification number.
        /// </summary>
        public int TagNumber
        {
            get { return Convert.ToInt32(TagData["tag_unique_id"]); }
            private set { TagData["tag_unique_id"] = value; }
        }

        /// <summary>
        ///     The full information profile of the current tag.  Controls all the tag properties.
        /// </summary>
        public Dictionary<string, object> TagData { get; set; }

        #endregion

        // Tag initializer
        public SprTag()
        {
            // Link the parent application
            Application = SprApplication.ActiveApplication;

            // Create a new data dictionary from the template
            TagData = SprUtilities.TagTemplate;

            // Set the tag to the next available tag number
            TagNumber = Application.NextTag;

            // Add the default data values
            BackgroundColor = SprUtilities.From0Bgr(12632319);
            LeaderColor = SprUtilities.From0Bgr(12632319);
        }
    }
}