# Outlook Contacts to .VCF
This connects to Outlook instance running on the computer and exports all contacts in the main profile into a VCF file.

Microsoft removed all feasible ways to export contacts as a .vcf file in Office 365. So if you want to migrate away with contact storage, you have no way to export them. Other corpos provide tools to export from Outlook and import to them, but if you're trying to move to self-hosted CardDAV server (like radicale), you're SOL. Even if you go through the pain and drag all contacts from Contacts view into a new mail window, Outlook will attach all of them as .msg, not .vcf. You can also try using the share button, which allows export as .vcf. It will crash if you have many contacts, or it will export them with broken encoding, i.e. if you live in a 1251 codepage region, the .vcf file names will be named correctly, but their content will be broken. And even if you somehow manage to do all this, you'll end up with hundreds of VCF files, but all tools that read them (radicale, Thunderbird, etc.) only allow importing them one by one. This is because VCF format (RFC 6350, https://datatracker.ietf.org/doc/html/rfc6350#section-3) allows for multiple contacts to be stored in a single VCF file, so in other creators' logic, you'd just use that and still have 1 VCF file to import. This tool simply iterates over all contacts in your main Outlook profile and puts them all into 1 VCF file in the directory you ran it from. You can then import it anywhere you need.

# You are a dev
Pull this repo, open in VS, press F5, it'll build and run. The .vcf file is stored in the debug folder, it has a guid name.

# You are a user
Download the .exe file, place it on a PC with Outlook and into a folder where you want the .vcf contacts stored and run it. It'll export all contacts into a file named something like asdf-ghjkla-asdfghg-jklh.vcf and that's the file you need for Thunderbird or other contact-keeping software. Talk to your IT-buddy before running .exe off the internet.
