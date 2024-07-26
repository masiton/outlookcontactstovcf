using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace OutlookContactsExtractor
{
    public class Extractor
    {
        public List<Contact> GetContacts()
        {
            var contacts = new List<Contact>();

            // https://stackoverflow.com/questions/17010963/how-to-get-the-contacts-of-outlook-to-c
            Microsoft.Office.Interop.Outlook.Items OutlookItems;
            Microsoft.Office.Interop.Outlook.Application outlookObj;
            MAPIFolder Folder_Contacts;

            outlookObj = new Microsoft.Office.Interop.Outlook.Application();
            Folder_Contacts = (MAPIFolder)outlookObj.Session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
            OutlookItems = Folder_Contacts.Items;

            for (int i = 0; i < OutlookItems.Count; i++)
            {
                Microsoft.Office.Interop.Outlook.ContactItem contact = (Microsoft.Office.Interop.Outlook.ContactItem)OutlookItems[i + 1];
                var c = new Contact()
                {
                    Name = contact.FirstName,
                    Surname = contact.LastName,
                    Phones = GetPhones(contact),
                    Emails = GetEmails(contact),
                    Company = contact.CompanyName,
                };

                contacts.Add(c);
            }

            return contacts;
        }

        private string[] GetPhones(ContactItem contact)
        {
            var phones = new List<string>();

            AddIfNotNull(contact.BusinessTelephoneNumber);
            AddIfNotNull(contact.HomeTelephoneNumber);
            AddIfNotNull(contact.MobileTelephoneNumber);
            AddIfNotNull(contact.OtherTelephoneNumber);

            return phones.ToArray();

            void AddIfNotNull(string phone)
            {
                if (string.IsNullOrWhiteSpace(phone) == false)
                {
                    phones.Add(phone);
                }
            }
        }

        private string[] GetEmails(ContactItem contact)
        {
            var phones = new List<string>();

            AddIfNotNull(contact.Email1Address);
            AddIfNotNull(contact.Email2Address);
            AddIfNotNull(contact.Email3Address);

            return phones.ToArray();

            void AddIfNotNull(string phone)
            {
                if (string.IsNullOrWhiteSpace(phone) == false)
                {
                    phones.Add(phone);
                }
            }
        }
    }
}
