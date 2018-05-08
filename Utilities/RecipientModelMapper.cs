using OutlookRuleMgr.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookRuleMgr.Utilities
{
    public class RecipientModelMapper
    {
        public Recipient MapToModel(Outlook.Recipient r)
            => new Recipient {Name = r.Name, EmailAddress = r.GetEmailAddress()};

        public void AddToOutlook(Recipient recipientModel, Outlook.Recipients recipients)
        {
            if (recipientModel.EmailAddress != null)
            {
                var name = recipientModel.Name != recipientModel.EmailAddress
                    ? $"{recipientModel.Name} <{recipientModel.EmailAddress}>"
                    : recipientModel.Name;
                recipients.Add(name).Resolve();
            }
            else
            {
                recipients.Add(recipientModel.Name);
            }
        }
    }
}
