using Library.Tools.Attributes;
using Library.Tools.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

/// XMind API for C#
/// ----------------
///
/// API for creating XMind workbook data files.
///
/// Note: This code is part of a CodePlex project: http://XMindAPI.codeplex.com
///
/// CHANGE LOG:
/// ==================================================================================================================
/// #       AUTHOR                  DATE          DESCRIPTION
/// ------  ----------------------  ------------  --------------------------------------------------------------------
/// 1.      Rudi Rademeyer          2011/06/04    Create.
/// 2.      Rudi Rademeyer          2012/03/08    Add functionality to read existing XMind workbook files.
/// 3.      Rudi Rademeyer          2012/10/28    Eliminated dependency on Java libraries for creating and opening zip
///                                               files by implementing ZipStorer from zipstorer.codeplex.com.
/// 4.      Julien Malbault         20106/08/04   Possibilité d'ajouter des images
///
/// ==================================================================================================================
namespace Xmind.Helper
{
    /// <summary>
    /// XMindStructure defines the different types of diagrams that can be drawn. Its implemented through the AddCentralTopic() method.
    /// </summary>
    public enum XMindStructure
    {
        [Description("org.xmind.ui.fishbone.rightHeaded")]
        FishboneRightHeaded,

        [Description("org.xmind.ui.fishbone.leftHeaded")]
        FishboneLeftHeaded,

        [Description("org.xmind.ui.spreadsheet")]
        SpreadSheet,

        [Description("org.xmind.ui.map")]
        Map,

        [Description("org.xmind.ui.map.clockwise")]
        MapClockwise,

        [Description("org.xmind.ui.map.anticlockwise")]
        MapAntiClockwise,

        [Description("org.xmind.ui.org-chart.down")]
        OrgChartDown,

        [Description("org.xmind.ui.org-chart.up")]
        OrgChartUp,

        [Description("org.xmind.ui.tree.left")]
        TreeLeft,

        [Description("org.xmind.ui.tree.right")]
        TreeRight,

        [Description("org.xmind.ui.logic.right")]
        LogicRight,

        [Description("org.xmind.ui.logic.left")]
        LogicLeft
    }

    /// <summary>
    /// XMindMarkers define optional markers that can be added to topics. Markers are displayed right in front of the topic title.
    /// Refer AddMarker() method.
    /// </summary>
    public enum XMindMarkers
    {
        // Task markers:
        [Description("task-start")]
        TaskStart,

        [Description("task-quarter")]
        TaskQuarter,

        [Description("task-half")]
        TaskHalf,

        [Description("task-3quar")]
        Task3Quarter,

        [Description("task-done")]
        TaskDone,

        [Description("task-paused")]
        TaskPaused,

        // Priority markers:
        [Description("priority-1")]
        Priority1,

        [Description("priority-2")]
        Priority2,

        [Description("priority-3")]
        Priority3,

        [Description("priority-4")]
        Priority4,

        [Description("priority-5")]
        Priority5,

        [Description("priority-6")]
        Priority6,

        // Smiley markers:
        [Description("smiley-smile")]
        SmieySmile,

        [Description("smiley-laugh")]
        SmileyLaugh,

        [Description("smiley-angry")]
        SmileyAngry,

        [Description("smiley-cry")]
        SmileyCry,

        [Description("smiley-surprise")]
        SmileySurprise,

        [Description("smiley-boring")]
        SmileyBoring,

        // Flag markers:
        [Description("flag-green")]
        FlagGreen,

        [Description("flag-red")]
        FlagRed,

        [Description("flag-orange")]
        FlagOrange,

        [Description("flag-purple")]
        FlagPurple,

        [Description("flag-blue")]
        FlagBlue,

        [Description("flag-black")]
        FlagBlack,

        // Star markers:
        [Description("star-green")]
        StarGreen,

        [Description("star-red")]
        StarRed,

        [Description("star-yellow")]
        StarYellow,

        [Description("star-purple")]
        StarPurple,

        [Description("star-blue")]
        StarBlue,

        [Description("star-gray")]
        StarGray,

        // Half Star markers:
        [Description("half-star-green")]
        HalfStarGreen,

        [Description("half-star-red")]
        HalfStarRed,

        [Description("half-star-yellow")]
        HalfStarYellow,

        [Description("half-star-purple")]
        HalfStarPurple,

        [Description("half-star-blue")]
        HalfStarBlue,

        [Description("half-star-gray")]
        HalfStarGray,

        // Other markers:
        [Description("other-calendar")]
        Caledar,

        [Description("other-email")]
        Email,

        [Description("other-phone")]
        Phone,

        [Description("other-phone")]
        Phone2,

        [Description("other-fax")]
        Fax,

        [Description("other-people")]
        People,

        [Description("other-people2")]
        People2,

        [Description("other-clock")]
        Clock,

        [Description("other-coffee-cup")]
        CoffeeCup,

        [Description("other-question")]
        Question,

        [Description("other-exclam")]
        ExclamationMark,

        [Description("other-lightbulb")]
        LightBulb,

        [Description("other-businesscard")]
        BusinessCard,

        [Description("other-social")]
        Social,

        [Description("other-chat")]
        Chat,

        [Description("other-note")]
        Note,

        [Description("other-lock")]
        Lock,

        [Description("other-unlock")]
        Unlock,

        [Description("other-yes")]
        Yes,

        [Description("other-no")]
        No,

        [Description("other-bomb")]
        Bomb
    }

    public enum XMindTopicShape
    {
        [Description("org.xmind.topicShape.roundedRect")]
        roundedRect,

        [Description("org.xmind.topicShape.underline")]
        underline,
    }

    public enum XMindBoundaryShape
    {
        [Description("org.xmind.boundaryShape.roundedPolygon")]
        roundedPolygon,
    }

    public enum XmindAlign
    {
        left,
        right,
        top,
        bottom,
    }

    public class XMindSheet
    {
        #region Public PROPERTIES

        public string ID { get; private set; }
        public string Name { get; private set; }

        public List<XMindTopic> TopicFlatList { get; private set; }
        public List<XMindTopic> Topics { get; private set; }

        #endregion

        #region Public METHODS

        public new string ToString()
        {
            return Name;
        }

        #endregion

        #region Internal CONSTRUCTORS

        internal XMindSheet(string sheetId, string sheetName)
            : this()
        {
            ID = sheetId;
            Name = sheetName;
        }

        #endregion

        #region Private CONSTRUCTORS

        private XMindSheet()
        {
            TopicFlatList = new List<XMindTopic>();
            Topics = new List<XMindTopic>();
        }

        #endregion
    }

    public class XMindTopic
    {
        #region Public PROPERTIES

        public string ID { get; private set; }
        public string Name { get; private set; }

        public XMindTopic Parent { get; private set; }

        public List<XMindTopic> Topics { get; private set; }

        #endregion

        #region Internal CONSTRUCTORS

        internal XMindTopic(XMindTopic parent, string id, string name)
            : this()
        {
            Parent = parent;
            ID = id;
            Name = name;
        }

        #endregion

        #region Private CONSTRUCTORS

        private XMindTopic()
        {
            Topics = new List<XMindTopic>();
        }

        #endregion
    }

    public class XMindStyle
    {
        #region Public PROPERTIES

        public string ID { get; private set; }
        public string BorderLineColor { get; private set; }
        public string BorderLineWidth { get; private set; }
        public string FontFamily { get; private set; }
        public string FontSize { get; private set; }
        public string FontStyle { get; private set; }
        public string ShapeClass { get; private set; }
        public string Fill { get; private set; }

        #endregion
    }

    /// <summary>
    /// XMindWorkBook encapsulates an XMind workbook and methods for performing actions on workbook content.
    /// </summary>
    public class XMindWorkBook
    {
        #region Public CONSTRUCTORS

        /// <summary>
        /// Creates a new XMind workbook. If the workbook exists it will be overwritten with a new empty workbook.
        /// </summary>
        /// <param name="fileName">XMind workbook file to create</param>
        public XMindWorkBook(string fileName)
            : this(fileName, false)
        {
        }

        /// <summary>
        /// Creates a new XMind workbook if loadContent is false, otherwise the file content will be loaded.
        /// </summary>
        /// <param name="fileName">XMind file to create / load</param>
        /// <param name="loadContent">If true, the current data from the file will be loaded, otherwise an empty workbook will be created.</param>
        public XMindWorkBook(string fileName, bool loadContent)
        {
            _fileName = fileName;
            _tempPath = Path.GetTempPath() + Guid.NewGuid() + "\\";
            Directory.CreateDirectory(_tempPath);

            _attachmentFolderPath = _tempPath + @"attachments\";
            Directory.CreateDirectory(_attachmentFolderPath);

            _defaultStylesNS = XNamespace.Get("urn:xmind:xmap:xmlns:style:2.0");
            _defaultContentNS = XNamespace.Get("urn:xmind:xmap:xmlns:content:2.0");
            _defaultManifestNS = XNamespace.Get("urn:xmind:xmap:xmlns:manifest:1.0");
            _defaultMetaNS = XNamespace.Get("urn:xmind:xmap:xmlns:meta:2.0");
            _xlinkNS = XNamespace.Get("http://www.w3.org/1999/xlink");
            _xhtmlNS = XNamespace.Get("http://www.w3.org/1999/xhtml");

            if (loadContent)
            {
                Load();
            }
            else
            {
                CreateDefaultMetaFile();
                CreateDefaultContentFile();
                CreateDefaultStylesFile();
            }
        }

        #endregion

        #region Public METHODS

        public List<XMindSheet> GetSheetInfo()
        {
            List<XMindSheet> lst = new List<XMindSheet>();

            foreach (XElement el in GetSheets())
            {
                string sheetId = GetAttribValue(el, "id");

                string centralTopicId = el.Descendants().Where(w1 => w1.Name.ToString().EndsWith("topic") && w1.Parent.Name.ToString().EndsWith("sheet"))
                    .Where(w3 => GetAttribValue(w3.Parent, "id") == sheetId)
                    .Select(s => GetAttribValue(s, "id")).First();

                string centralTopicTitle = GetTopic(centralTopicId).Descendants()
                    .Where(w2 => w2.Name.ToString().EndsWith("title")).Select(s => s.Value).First();

                XMindSheet xmSheet = new XMindSheet(sheetId,
                    el.Descendants()
                    .Where(w2 => w2.Name.ToString().EndsWith("title") && w2.Parent.Name.ToString().EndsWith("sheet"))
                    .Where(w3 => GetAttribValue(w3.Parent, "id") == sheetId)
                    .Select(s => s.Value).First());

                XMindTopic xmTopic = new XMindTopic(null, centralTopicId, centralTopicTitle);
                xmSheet.Topics.Add(xmTopic);

                xmSheet.TopicFlatList.Add(xmTopic);
                GetTopicsRecursively(GetTopic(centralTopicId), xmSheet, xmTopic);

                lst.Add(xmSheet);
            }

            return lst;
        }

        /// <summary>
        /// Add a new sheet to the workbook.
        /// </summary>
        /// <param name="sheetName">Name of the new sheet (sheet title)</param>
        /// <returns>New seet id</returns>
        public string AddSheet(string sheetName)
        {
            string sheetId = NewId();

            _contentData.Root.Add(
                new XElement(_defaultContentNS + "sheet",
                    new XAttribute("id", sheetId),
                    new XAttribute("timestamp", GetTimeStamp()),
                    new XElement(_defaultContentNS + "title", sheetName)
                    ));

            return sheetId;
        }

        /// <summary>
        /// Get a list of sheet id's of all sheets matching the specified sheet title. Note: It is possible to have more than
        /// one sheet in the same workbook with the same title.
        /// </summary>
        /// <param name="title">Sheet title to search for</param>
        /// <returns>List of sheet id's</returns>
        public List<string> GetSheetIdsByTitle(string title)
        {
            List<string> sheetsFound = new List<string>();

            foreach (XElement sheet in GetSheets())
            {
                sheetsFound.AddRange(sheet.Descendants()
                    .Where(w2 => w2.Name.ToString().EndsWith("title") && w2.Value == title && w2.Parent.Name.ToString().EndsWith("sheet"))
                    .Select(s => GetAttribValue(s.Parent, "id")).ToList());
            }

            return sheetsFound;
        }

        /// <summary>
        /// Add a new central topic to the specified sheet. A sheet must have one (and only one) central topic.
        /// </summary>
        /// <param name="sheetId">Sheet to add central topic to</param>
        /// <param name="topicName">Name of the central topic</param>
        /// <param name="structure">Type of diagram structure. Refer XMindStructure enum.</param>
        /// <returns>Id of newly created central topic</returns>
        public string AddCentralTopic(string sheetId, string topicName, XMindStructure structure)
        {
            XElement sheet = GetSheet(sheetId);

            if (sheet == null)
            {
                throw new InvalidOperationException("Sheet not found!");
            }

            if (GetTopics(sheet).Count() > 0)
            {
                throw new InvalidOperationException("Sheet can have only one central topic!");
            }

            string topicId = NewId();

            var enumField = structure.GetType().GetFields().Where(field => field.Name == structure.ToString()).FirstOrDefault();
            DescriptionAttribute[] a = (DescriptionAttribute[])enumField.GetCustomAttributes(typeof(DescriptionAttribute), false);

            sheet.Add(
                new XElement(_defaultContentNS + "topic",
                    new XAttribute("id", topicId),
                    new XAttribute("structure-class", a[0].Description),
                    new XAttribute("timestamp", GetTimeStamp()),
                    new XElement(_defaultContentNS + "title", topicName)
                    ));

            return topicId;
        }

        /// <summary>
        /// Add a topic to either a central topic or another topic. Diagram structure can be specified, if set to null
        /// parent structure will be inherited.
        /// </summary>
        /// <param name="parentId">Id of parent topic</param>
        /// <param name="topicName">New topic title</param>
        /// <param name="structure">Type of diagram structure. Refer XMindStructure enum.</param>
        /// <returns>Newly created topic id</returns>
        public string AddTopic(string parentId, string topicName, XMindStructure? structure)
        {
            XElement parent = GetTopic(parentId);

            if (parent == null)
            {
                throw new InvalidOperationException("Topic not found!");
            }

            // Get topic children tag, if not exist create:
            XElement children = parent.Descendants().Where(w => w.Name.ToString().EndsWith("children")).FirstOrDefault();

            if (children == null)
            {
                children = new XElement(_defaultContentNS + "children");
                parent.Add(children);
            }

            // Get topics tag, if not exists create:
            XElement topics = children.Descendants().Where(w => w.Name.ToString().EndsWith("topics")).FirstOrDefault();

            if (topics == null)
            {
                topics = new XElement(_defaultContentNS + "topics",
                    new XAttribute("type", "attached"));
                children.Add(topics);
            }

            // Add new topic to topics element:
            string topicId = NewId();
            XElement topicElement = new XElement(_defaultContentNS + "topic",
                    new XAttribute("id", topicId),
                    new XAttribute("timestamp", GetTimeStamp()),
                    new XElement(_defaultContentNS + "title", topicName)
                    );

            if (structure != null)
            {
                var enumField = structure.GetType().GetFields().Where(field => field.Name == structure.ToString()).FirstOrDefault();
                DescriptionAttribute[] a = (DescriptionAttribute[])enumField.GetCustomAttributes(typeof(DescriptionAttribute), false);

                topicElement.Add(new XAttribute("structure-class", a[0].Description));
            }

            topics.Add(topicElement);

            return topicId;
        }

        /// <summary>
        /// Add a topic to either a central topic or another topic. Diagram structure can be specified, if set to null
        /// parent structure will be inherited.
        /// </summary>
        /// <param name="parentId">Id of parent topic</param>
        /// <param name="topicName">New topic title</param>
        /// <param name="structure">Type of diagram structure. Refer XMindStructure enum.</param>
        /// <returns>Newly created topic id</returns>
        public string AddTopicWithStyle(string parentId, string topicName, string styleId)
        {
            XElement parent = GetTopic(parentId);

            if (parent == null)
            {
                throw new InvalidOperationException("Topic not found!");
            }

            // Get topic children tag, if not exist create:
            XElement children = parent.Descendants().Where(w => w.Name.ToString().EndsWith("children")).FirstOrDefault();

            if (children == null)
            {
                children = new XElement(_defaultContentNS + "children");
                parent.Add(children);
            }

            // Get topics tag, if not exists create:
            XElement topics = children.Descendants().Where(w => w.Name.ToString().EndsWith("topics")).FirstOrDefault();

            if (topics == null)
            {
                topics = new XElement(_defaultContentNS + "topics",
                    new XAttribute("type", "attached"));
                children.Add(topics);
            }

            // Add new topic to topics element:
            string topicId = NewId();
            XElement topicElement = new XElement(_defaultContentNS + "topic",
                    new XAttribute("id", topicId),
                    new XAttribute("style-id", styleId),
                    new XAttribute("timestamp", GetTimeStamp()),
                    new XElement(_defaultContentNS + "title", topicName)
                    );

            topics.Add(topicElement);

            return topicId;
        }

        /// <summary>
        /// Ajout d'un lien
        /// </summary>
        /// <param name="topicId"></param>
        /// <param name="link"></param>
        public void AddLink(string topicId, string link)
        {
            XElement topic = GetTopic(topicId);

            if (topic == null)
                throw new InvalidOperationException("Topic not found!");

            var prefixe = string.Empty;

            var uri = new Uri(link);
            if (uri.IsFile)
                prefixe = "file://";
            else if (uri.IsHTTP())
                prefixe = "";
            else
                throw new Exception("Les autres types de lien hypertext ne sont pas supporté.");

            topic.Add(new XAttribute(_xlinkNS + "href", prefixe + link));
        }

        /// <summary>
        /// Add a topic to either a central topic or another topic. Parent diagram structure will be inherited.
        /// </summary>
        /// <param name="parentId">Id of parent topic</param>
        /// <param name="topicName">New topic title</param>
        /// <returns>Newly created topic id</returns>
        public string AddTopic(string parentId, string topicName)
        {
            return AddTopic(parentId, topicName, null);
        }

        /// <summary>
        /// Add a label to the specified topic.
        /// </summary>
        /// <param name="topicId">Id of topic to add label to</param>
        /// <param name="labelText">Label text</param>
        public void AddLabel(string topicId, string labelText)
        {
            XElement topic = GetTopic(topicId);

            if (topic == null)
            {
                throw new InvalidOperationException("Topic not found!");
            }

            // Get topic labels tag, if not exist create:
            XElement labels = topic.Descendants().Where(w => w.Name.ToString().EndsWith("labels")).FirstOrDefault();

            if (labels == null)
            {
                labels = new XElement(_defaultContentNS + "labels");
                topic.Add(labels);
            }

            // Get topic label tag, if not exist create:
            XElement label = labels.Descendants().Where(w => w.Name.ToString().EndsWith("label")).FirstOrDefault();

            if (label == null)
            {
                label = new XElement(_defaultContentNS + "label");
                labels.Add(label);
            }

            label.Value = labelText;
        }

        /// <summary>
        /// Add a picture to the specified topic
        /// </summary>
        /// <param name="topicId"></param>
        /// <param name="imagePath"></param>
        public void AddPicture(string topicId, string imagePath, XmindAlign align )
        {
            XElement topic = GetTopic(topicId);

            if (topic == null)
                throw new InvalidOperationException("Topic not found!");

            //store picture in attachments
            var attachmentFilePath = _attachmentFolderPath + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 24) + Path.GetExtension(imagePath);
            File.Copy(imagePath, attachmentFilePath);

            //Ajout img
            XElement img = topic.Descendants().Where(w => w.Name.ToString().EndsWith("img")).FirstOrDefault();

            //Align
            string alignString = string.Empty;

            if (align == XmindAlign.left)
                alignString = "left";
            else if (align == XmindAlign.right)
                alignString = "right";
            else if (align == XmindAlign.top)
                alignString = "top";
            else if (align == XmindAlign.bottom)
                alignString = "bottom";
            else
                throw new NotSupportedException(align.ToStringWithEnumName());

            if (img == null)
            {
                img = new XElement(_xhtmlNS + "img",
                    new XAttribute("align", alignString),
                    new XAttribute(_xhtmlNS + "src", "xap:attachments/" + Path.GetFileName(attachmentFilePath)));
                topic.Add(img);
            }
        }

        /// <summary>
        /// Add a marker to an existing topic. Refer XMindMarkers enum for available markers.
        /// </summary>
        /// <param name="topicId">Id of topic to add marker to</param>
        /// <param name="marker">Marker type. Refer XMindMarkers enum</param>
        public void AddMarker(string topicId, XMindMarkers marker)
        {
            XElement topic = GetTopic(topicId);

            if (topic == null)
            {
                throw new InvalidOperationException("Topic not found!");
            }

            // Get topic marker-refs tag, if not exist create:
            XElement marker_refs = topic.Descendants().Where(w => w.Name.ToString().EndsWith("marker-refs")).FirstOrDefault();

            if (marker_refs == null)
            {
                marker_refs = new XElement(_defaultContentNS + "marker-refs");
                topic.Add(marker_refs);
            }

            // Get topic marker_ref tag, if not exist create:
            XElement marker_ref = marker_refs.Descendants().Where(w => w.Name.ToString().EndsWith("marker-ref")).FirstOrDefault();

            if (marker_ref == null)
            {
                marker_ref = new XElement(_defaultContentNS + "marker-ref");
                marker_refs.Add(marker_ref);
            }

            XAttribute att = marker_ref.Attributes().Where(w => w.Name == "marker-id").FirstOrDefault();

            if (att != null)
            {
                marker_ref.Attributes("marker-id").Remove();
            }

            var enumField = marker.GetType().GetFields().Where(field => field.Name == marker.ToString()).FirstOrDefault();
            DescriptionAttribute[] a = (DescriptionAttribute[])enumField.GetCustomAttributes(typeof(DescriptionAttribute), false);

            marker_ref.Add(new XAttribute("marker-id", a[0].Description));
        }

        /// <summary>
        /// Add a link from one topic to another.
        /// </summary>
        /// <param name="topicId">Id of topic to contain the link</param>
        /// <param name="linkToTopicId">Id of topic to link to</param>
        public void AddTopicLink(string topicId, string linkToTopicId)
        {
            XElement topic = GetTopic(topicId);

            if (topic == null)
                throw new InvalidOperationException("Topic not found!");

            if (GetTopic(linkToTopicId) == null)
                throw new InvalidOperationException("Link to topic not found!");

            XAttribute att = topic.Attributes().Where(w => w.Name == _xlinkNS + "href").FirstOrDefault();

            if (att != null)
                topic.Attributes(_xlinkNS + "href").Remove();

            topic.Add(new XAttribute(_xlinkNS + "href", "xmind:#" + linkToTopicId));
        }

        /// <summary>
        /// Collapse the child structure for the specified topic.
        /// </summary>
        /// <param name="topicId">Topic to collapse child structure</param>
        public void CollapseChildren(string topicId)
        {
            XElement topic = GetTopic(topicId);

            if (topic == null)
            {
                throw new InvalidOperationException("Topic not found!");
            }

            XAttribute att = topic.Attributes().Where(w => w.Name == "branch").FirstOrDefault();

            if (att != null)
            {
                topic.Attributes("branch").Remove();
            }

            topic.Add(new XAttribute("branch", "folded"));
        }

        /// <summary>
        /// Change the title of a topic.
        /// </summary>
        /// <param name="topicId">Id of topic</param>
        /// <param name="newTitle">New title</param>
        public void EditTopicTitle(string topicId, string newTitle)
        {
            XElement topic = GetTopic(topicId);

            if (topic != null)
            {
                XElement titleElement = topic.Descendants().Where(w => w.Name.ToString().EndsWith("title")).First();
                titleElement.Value = newTitle;
            }
        }

        /// <summary>
        /// Get the title text for the specified topic id.
        /// </summary>
        /// <param name="topicId">Topic id</param>
        /// <returns>Topic title</returns>
        public string GetTopicTitle(string topicId)
        {
            string title = null;
            XElement topic = GetTopic(topicId);

            if (topic != null)
            {
                title = topic.Descendants().Where(w => w.Name.ToString().EndsWith("title")).Select(s => s.Value).First();
            }

            return title;
        }

        /// <summary>
        /// Get a list of topic id's where the title matches the suppled title. All sheets will be searched.
        /// </summary>
        /// <param name="title">Topic title to search for</param>
        /// <returns>List of topic titles found</returns>
        public List<string> GetTopicIdsByTitle(string title)
        {
            List<string> topicsFound = new List<string>();

            foreach (XElement sheet in GetSheets())
            {
                topicsFound.AddRange(GetTopicIdsByTitle(GetAttribValue(sheet, "id"), title));
            }

            return topicsFound;
        }

        /// <summary>
        /// Get a list of topic id's where the title matches the suppled title. Only the specified sheet will be searched.
        /// </summary>
        /// <param name="sheetId">Sheet to search in</param>
        /// <param name="title">Topic title to search for</param>
        /// <returns>List of topic titles found</returns>
        public List<string> GetTopicIdsByTitle(string sheetId, string title)
        {
            List<string> topicsFound = new List<string>();

            XElement sheet = GetSheet(sheetId);

            if (sheet != null)
            {
                topicsFound.AddRange(sheet.Descendants().Where(w1 => w1.Name.ToString().EndsWith("topic"))
                    .Descendants().Where(w2 => w2.Name.ToString().EndsWith("title") && w2.Value == title).Select(s => GetAttribValue(s.Parent, "id")).ToList());
            }

            return topicsFound;
        }

        /// <summary>
        /// Get a list of topic id's that contains the specified user tag where the user tag value mathes the specified value.
        /// Also see method AddUserTag().
        /// </summary>
        /// <param name="tagName">User tag to search</param>
        /// <param name="searchValue">User tag value to match</param>
        /// <returns>List of topic id's where the user tag/value matches</returns>
        public List<string> GetTopicIdsByUserTagValue(string tagName, string searchValue)
        {
            List<string> topicsFound = new List<string>();

            foreach (XElement sheet in GetSheets())
            {
                foreach (XElement topic in GetTopics(sheet))
                {
                    string topicId = GetAttribValue(topic, "id");
                    if (GetUserTagValues(topicId, tagName).Contains(searchValue))
                    {
                        topicsFound.Add(topicId);
                    }
                }
            }

            return topicsFound;
        }

        /// <summary>
        /// Add style to be able to use it in a topic
        /// </summary>
        public string AddStyleTopic(XMindTopicShape iShapeClass, System.Drawing.Color iFill)
        {
            string styleId = NewId();
            var colorConvertor = new ColorConverter();

            var xElement = GetStyles();

            xElement.Add(
                new XElement(_defaultStylesNS + "style",
                    new XAttribute("id", styleId),
                    new XAttribute("type", "topic"),
                    new XElement(_defaultStylesNS + "topic-properties",
                        new XAttribute("fill", iFill.HexConverter()),
                        new XAttribute("shape-class", iShapeClass.GetDescription()))
                    ));

            return styleId;
        }

        /// <summary>
        /// Add a user tag to the specified sheet or topic.
        /// </summary>
        /// <param name="itemId">Sheet or topic id to add user tag to</param>
        /// <param name="tagName">User tag name</param>
        /// <param name="tagValue">User tag value</param>
        public void AddUserTag(string itemId, string tagName, string tagValue)
        {
            // Check if itemId is a sheet:
            XElement item = GetSheet(itemId);

            // If not a sheet, check if itemid is a topic:
            if (item == null)
            {
                item = GetTopic(itemId);
            }

            if (item == null)
            {
                throw new InvalidOperationException("Topic/Sheet not found!");
            }

            // Get user tags, if not exist create:
            XElement userTags = item.Descendants().Where(w => w.Name.ToString().EndsWith("UserTags")).FirstOrDefault();

            if (userTags == null)
            {
                userTags = new XElement(_defaultContentNS + "UserTags");
                item.Add(userTags);
            }

            // Get the named user tag, if not exist create:
            XElement userTag = userTags.Descendants()
                .Where(w => w.Name.ToString().EndsWith("UserTag") && GetAttribValue(w, "TagName") == tagName).FirstOrDefault();

            if (userTag == null)
            {
                userTag = new XElement(_defaultContentNS + "UserTag",
                    new XAttribute("TagName", tagName),
                    new XAttribute("TagValue", ""));
                userTags.Add(userTag);
            }

            userTag.SetAttributeValue(XName.Get("TagValue"), tagValue);
        }

        /// <summary>
        /// Get the values of a specified user tag from the specified sheet/topic.
        /// </summary>
        /// <param name="itemId">Id of sheet or topic to search</param>
        /// <param name="tagName">User tag name to search for</param>
        /// <returns>List of user tag values that was found</returns>
        public List<string> GetUserTagValues(string itemId, string tagName)
        {
            List<string> tagValues = new List<string>();

            // Check if itemId is a sheet:
            XElement item = GetSheet(itemId);

            // If not a sheet, check if itemid is a topic:
            if (item == null)
            {
                item = GetTopic(itemId);
            }

            if (item == null)
            {
                return tagValues;
            }

            // Get user tags:
            XElement userTags = item.Descendants().Where(w => w.Name.ToString().EndsWith("UserTags")).FirstOrDefault();

            if (userTags == null)
            {
                return tagValues;
            }

            // Get the named user tag:
            foreach (XElement userTag in userTags.Descendants()
                .Where(w => w.Name.ToString().EndsWith("UserTag") && GetAttribValue(w, "TagName") == tagName))
            {
                tagValues.Add(GetAttribValue(userTag, "TagValue"));
            }

            return tagValues;
        }

        /// <summary>
        /// Save the current XMind workbook file to disk.
        /// </summary>
        public void Save()
        {
            if (_fileName == null)
            { throw new InvalidOperationException("Nothing to save!"); }

            Directory.CreateDirectory(_tempPath + "META-INF");
            Directory.CreateDirectory(_tempPath + "Thumbnails");
            CreateManifestFile();
            File.WriteAllText(_tempPath + "META-INF\\manifest.xml", _manifestData.ToString());
            File.WriteAllText(_tempPath + "meta.xml", _metaData.ToString());
            File.WriteAllText(_tempPath + "content.xml", _contentData.ToString());
            File.WriteAllText(_tempPath + "styles.xml", _stylesData.ToString());

            System.IO.Compression.ZipFile.CreateFromDirectory(_tempPath, _fileName);

            Directory.Delete(_tempPath, true);
        }

        #endregion

        #region Private FIELDS

        private string _fileName = null;
        private string _tempPath = null;
        private string _attachmentFolderPath = null;
        private XDocument _manifestData = null;
        private XDocument _metaData = null;
        private XDocument _contentData = null;
        private XDocument _stylesData = null;
        private XNamespace _defaultStylesNS = null;
        private XNamespace _defaultContentNS = null;
        private XNamespace _defaultManifestNS = null;
        private XNamespace _defaultMetaNS = null;
        private XNamespace _xhtmlNS = null;
        private XNamespace _xlinkNS = null;

        #endregion

        #region Private METHODS

        /// <summary>
        /// Helper method to build nested topic structure used by public method GetSheetInfo().
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="xmSheet"></param>
        /// <param name="xmTopic"></param>
        private void GetTopicsRecursively(XElement parent, XMindSheet xmSheet, XMindTopic xmTopic)
        {
            foreach (XElement nextLevelTopic in parent.Descendants().Where(w1 => w1.Name.ToString().EndsWith("topic")
                && GetAttribValue(w1.Parent.Parent.Parent, "id") == GetAttribValue(parent, "id")))
            {
                string topicId = GetAttribValue(nextLevelTopic, "id");
                XMindTopic nextXmTopic = new XMindTopic(xmTopic, topicId, GetTopicTitle(topicId));

                xmSheet.TopicFlatList.Add(nextXmTopic);
                xmTopic.Topics.Add(nextXmTopic);
                GetTopicsRecursively(nextLevelTopic, xmSheet, nextXmTopic);
            }
        }

        private XElement GetSheet(string sheetId)
        {
            return GetSheets().Where(w => GetAttribValue(w, "id") == sheetId).FirstOrDefault();
        }

        private List<XElement> GetSheets()
        {
            return _contentData.Root.Elements().Where(w => w.Name.ToString().EndsWith("sheet")).ToList();
        }

        private XElement GetTopic(string topicId)
        {
            XElement topic = null;

            foreach (XElement sheet in GetSheets())
            {
                topic = GetTopics(sheet).Where(w => GetAttribValue(w, "id") == topicId).FirstOrDefault();

                if (topic != null) break;
            }

            return topic;
        }

        private List<XElement> GetTopics(XElement sheet)
        {
            return sheet.Descendants().Where(w => w.Name.ToString().EndsWith("topic")).ToList();
        }

        private XElement GetStyles()
        {
            return _stylesData.Root.Elements().Single(w => w.Name.LocalName == "styles");
        }

        private string GetAttribValue(XElement el, string attributeName)
        {
            XAttribute att = el.Attributes(attributeName).FirstOrDefault();

            if (att == null)
            { return null; }
            else
            { return att.Value; }
        }

        private string NewId()
        {
            return Guid.NewGuid().ToString().Replace("-", "");
        }

        private string GetTimeStamp()
        {
            return DateTime.UtcNow.Ticks.ToString();
        }

        private void CreateDefaultMetaFile()
        {
            _metaData = new XDocument();

            _metaData.Declaration = new XDeclaration("1.0", "UTF-8", "no");

            _metaData.Add(
                new XElement(_defaultMetaNS + "meta",
                                new XAttribute("version", "2.0")));
        }

        private void CreateManifestFile()
        {
            _manifestData = new XDocument();

            _manifestData.Declaration = new XDeclaration("1.0", "UTF-8", "no");

            XElement rootEle = new XElement(_defaultManifestNS + "manifest");

            rootEle.Add(
               new XElement(_defaultManifestNS + "file-entry",
                   new XAttribute("full-path", "attachments/"),
                   new XAttribute("media-type", "")
               ));

            foreach (XElement element in _contentData.Descendants(_xhtmlNS + "img"))
            {
                var attribute = element.Attribute(_xhtmlNS + "src").Value.Replace("xap:", "");
                rootEle.Add(
               new XElement(_defaultManifestNS + "file-entry",
                   new XAttribute("full-path", attribute),
                   new XAttribute("media-type", "image/" + Path.GetExtension(attribute).Replace(".", ""))
               ));
            }

            rootEle.Add(
                new XElement(_defaultManifestNS + "file-entry",
                    new XAttribute("full-path", "content.xml"),
                    new XAttribute("media-type", "text/xml")
                ));

            rootEle.Add(
                new XElement(_defaultManifestNS + "file-entry",
                    new XAttribute("full-path", "styles.xml"),
                    new XAttribute("media-type", "text/xml")
                ));

            rootEle.Add(
                new XElement(_defaultManifestNS + "file-entry",
                    new XAttribute("full-path", "META-INF/"),
                    new XAttribute("media-type", "")
                ));

            rootEle.Add(
                new XElement(_defaultManifestNS + "file-entry",
                    new XAttribute("full-path", "META-INF/manifest.xml"),
                    new XAttribute("media-type", "text/xml")
                ));

            rootEle.Add(
                new XElement(_defaultManifestNS + "file-entry",
                    new XAttribute("full-path", "Thumbnails/"),
                    new XAttribute("media-type", "")
                ));

            _manifestData.Add(rootEle);
        }

        private void CreateDefaultContentFile()
        {
            _contentData = new XDocument();

            _contentData.Declaration = new XDeclaration("1.0", "UTF-8", "no");

            XNamespace ns2 = XNamespace.Get("http://www.w3.org/1999/XSL/Format");
            XNamespace ns3 = XNamespace.Get("http://www.w3.org/2000/svg");
            XNamespace ns4 = XNamespace.Get("http://www.w3.org/1999/xhtml");

            _contentData.Add(new XElement(_defaultContentNS + "xmap-content",
                new XAttribute(XNamespace.Xmlns + "fo", ns2),
                new XAttribute(XNamespace.Xmlns + "svg", ns3),
                new XAttribute(XNamespace.Xmlns + "xhtml", ns4),
                new XAttribute(XNamespace.Xmlns + "xlink", _xlinkNS),
                new XAttribute("version", "2.0")
                ));
        }

        private void CreateDefaultStylesFile()
        {
            _stylesData = new XDocument();

            _stylesData.Declaration = new XDeclaration("1.0", "UTF-8", "no");

            XNamespace ns2 = XNamespace.Get("http://www.w3.org/1999/XSL/Format");
            XNamespace ns3 = XNamespace.Get("http://www.w3.org/2000/svg");
            XNamespace ns4 = XNamespace.Get("http://www.w3.org/1999/xhtml");

            _stylesData.Add(new XElement(_defaultStylesNS + "xmap-styles",
                new XAttribute(XNamespace.Xmlns + "fo", ns2),
                new XAttribute(XNamespace.Xmlns + "svg", ns3),
                new XAttribute("version", "2.0")
                ));

            _stylesData.Root.Add(new XElement(_defaultStylesNS + "master-styles"));
            _stylesData.Root.Add(new XElement(_defaultStylesNS + "automatic-styles"));
            _stylesData.Root.Add(new XElement(_defaultStylesNS + "styles"));
        }

        private void Load()
        {
            if (_fileName == null)
            { throw new InvalidOperationException("No XMind file to load!"); }

            if (File.Exists(_fileName) == false)
            { throw new InvalidOperationException("XMind file does not exist!"); }

            FileInfo xMindFileInfo = new FileInfo(_fileName);

            if (xMindFileInfo.Extension.ToLower() != ".xmind")
            { throw new InvalidOperationException("XMind file extension expected!"); }

            String tempPath = Path.GetTempPath() + Guid.NewGuid() + "\\";

            System.IO.Compression.ZipFile.ExtractToDirectory(_fileName, tempPath);
        }

        #endregion
    }
}