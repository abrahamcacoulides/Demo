namespace Engine.Models
{
    public class DWG
    {
        public string JobNum { get; set; }
        public string PageID { get; set; }
        public string FileName { get; set; }
        public string PathToFile { get; set; }

        public DWG (string jobNum, string pageID, string fileName, string pathTofile)
        {
            JobNum = jobNum;
            PageID = pageID;
            FileName = fileName;
            PathToFile = pathTofile;
        }
    }
}
