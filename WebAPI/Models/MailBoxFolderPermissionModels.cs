namespace WebAPI.Models
{
    public class MailBoxFolderPermissionModels
    {
        public class GetMailBoxFolderPermission
        {
            public string UserPrincipalName { get; set; }
            public string DisplayName { get; set; }
            public string AliasName { get; set; }
            public object AccessRights { get; set; }
        }
        public class PostPutMailBoxFolderPermission
        {
            public string FolderPath { get; set; }
            public string UserPrincipalName { get; set; }
            public string[] AccessRights { get; set; }
        }

    }
}