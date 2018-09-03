using System.Collections.Generic;
using Engine.Models;

namespace Engine.Factories
{
    public static class VersionsFactory
    {
        private static List<AVersion> _listOfVersions;

        static VersionsFactory()
        {
            _listOfVersions = new List<AVersion>();

            _listOfVersions.Add(new AVersion(1, "AutoCAD R13", "\"C:\\r13\\win\\acad.exe\"", "C:\\r13\\win\\acad.exe", true));
            _listOfVersions.Add(new AVersion(2, "AutoCAD 14", "\"C:\\Program Files\\Autodesk\\AutoCAD 2014\\acad.exe\" ",
                "C:\\Program Files\\Autodesk\\AutoCAD 2014\\acad.exe", true));
            _listOfVersions.Add(new AVersion(3, "AutoCAD 17 R1", "\"C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe\" ",
                "C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe", true));
            _listOfVersions.Add(new AVersion(4, "AutoCAD 17", "\"C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe\" ",
                "C:\\Program Files\\Autodesk\\AutoCAD 2017\\acad.exe", false));
        }
    }
}
