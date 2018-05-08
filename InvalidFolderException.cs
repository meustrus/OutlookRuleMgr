using System;

namespace OutlookRuleMgr
{
    [Serializable]
    public class InvalidFolderException : Exception
    {
        public InvalidFolderException(string message) : base(message)
        {
        }
    }
}
