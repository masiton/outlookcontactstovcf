using FolkerKinzel.VCards;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookContactsExtractor
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var extractor = new Extractor();
            var contacts = extractor.GetContacts();

            var cards = new List<VCard>();

            foreach (var contact in contacts ) 
            {
                var card = new VCard();
                var builder = VCardBuilder.Create(card);

                contact.Phones.ToList().ForEach(x => builder.Phones.Add(x));
                contact.Emails.ToList().ForEach(x => builder.EMails.Add(x));
                builder.NameViews.Add(new[] { contact.Surname }, new[] {contact.Name});

                if(string.IsNullOrWhiteSpace(contact.Company) == false)
                {
                    builder.Organizations.Add(new FolkerKinzel.VCards.Models.Organization(contact.Company));
                }

                cards.Add(card);
            }

            var filename = $"{Guid.NewGuid()}.vcf";
            using (var stream = new FileStream(filename, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                Vcf.Serialize(cards, stream);
            }

        }
    }
}