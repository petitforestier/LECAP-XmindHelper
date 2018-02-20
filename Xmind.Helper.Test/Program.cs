using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xmind.Helper;

namespace XmindApi.ConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath = Library.Tools.IO.MyDirectory.GetSolutionDirectory() + ResultXmindFolder + Guid.NewGuid() +  ".xmind";

            // Create a new, empty workbook. If the workbook exists it will be overwritten:
            XMindWorkBook xwb = new XMindWorkBook(filePath);

            //Ajout des styles
            var underlineType = xwb.AddStyleTopic(XMindTopicShape.underline, System.Drawing.Color.LightGray);

            // Create a new sheet (at least one per workbook required):
            string sheetId = xwb.AddSheet("Vehicles");

            string centralTopicId = xwb.AddCentralTopic(sheetId, "Brands",XMindStructure.TreeRight);

            string mazdaTopicId = xwb.AddTopicWithStyle(centralTopicId, "Mazda", underlineType);
            string fordTopicId = xwb.AddTopic(centralTopicId, "Ford");
            string bmwTopicId = xwb.AddTopic(centralTopicId, "BMW");
            string nissanTopicId = xwb.AddTopic(centralTopicId, "Nissan", XMindStructure.TreeRight);

            string cx7TopicId = xwb.AddTopic(mazdaTopicId, "CX7");
            xwb.AddTopic(mazdaTopicId, "323");
            xwb.AddLink(mazdaTopicId, Library.Tools.IO.MyDirectory.GetSolutionDirectory() + DOC1);
            xwb.AddTopic(mazdaTopicId, "Mazda6");

            var imagePath = Library.Tools.IO.MyDirectory.GetSolutionDirectory() + IMG1;
            xwb.AddPicture(mazdaTopicId, imagePath);

            xwb.AddTopic(fordTopicId, "Bantam");
            xwb.AddTopic(fordTopicId, "Focus");
            xwb.AddTopic(fordTopicId, "Ranger");

            xwb.AddTopic(bmwTopicId, "3 series");
            xwb.AddTopic(bmwTopicId, "5 series");
            xwb.AddTopic(bmwTopicId, "7 series");

            xwb.AddTopic(nissanTopicId, "Nirvada");
            xwb.AddTopic(nissanTopicId, "Sentra");
            xwb.AddTopic(nissanTopicId, "Micra");

            xwb.AddLabel(cx7TopicId, "This is a SUV");

            xwb.AddMarker(bmwTopicId, XMindMarkers.FlagBlue);

            xwb.CollapseChildren(bmwTopicId);

            xwb.Save();
            Library.Tools.Debug.MyDebug.PrintInformation(filePath);
        }

        private const string ResultXmindFolder = @"_UnitTestFiles\";
        private const string IMG1 = @"_UnitTestFiles\img1.png";
        private const string IMG2 = @"_UnitTestFiles\img2.png";
        private const string DOC1 = @"_UnitTestFiles\test.docx";
    }
}
