using System;
using System.Management.Automation;

namespace WebAPI.Models
{
    public class MailBoxFolderModels
    {
        public class RootFolderModel
        {
            public string Name { get; set; }
            public Boolean HasSubfolders { get; set; }
        }

        public class SubFolderModel
        {
            public string Name { get; set; }
            public Boolean HasSubfolders { get; set; }
            public object FolderPath { get; set; }
        }
    }
}