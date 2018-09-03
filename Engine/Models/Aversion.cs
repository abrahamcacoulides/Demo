namespace Engine.Models
{
    public class AVersion
    {
        public int ID { get; set; }
        public string Version { get; set; }
        public string PathToVersion { get; set; }
        public string CleanPathToVersion { get; set; }
        public bool RequiresJobNum { get; set; }

        public AVersion(int id, string version, string pathToVersion, string cleanPathToVersion, bool requiresJobNum)
        {
            ID = id;
            Version = version;
            PathToVersion = pathToVersion;
            CleanPathToVersion = cleanPathToVersion;
            RequiresJobNum = requiresJobNum;
        }
    }
}
