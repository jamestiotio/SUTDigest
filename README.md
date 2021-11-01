# SUTDigest

SUTD Outlook Mail Classifier/Organizer

This tool will automatically create Outlook email rules, as well as the corresponding folders, to categorize SUTD emails. It can be run locally to apply both server-side and client-only rules.

Do take note that the Outlook rules will only apply to messages/emails received AFTER the rules are created. They will not be retroactively applied to past/old messages.

## Motivation

Because life is too short to:

- Read irrelevant emails that spam your inbox and drown out the important ones (and potentially tire or burn yourself out due to email fatigue)
- Fiddle around with Outlook's GUI client settings
- Attempt to import/export/share RWZ (rules) and PST (folder structure) files with empty message data successfully
- Convince SUTD IT to enable Microsoft Graph API for students to use freely

Simply put, we just want a double-clickable binary file that can setup, settle and resolve everything.

## Usage

> NOTE: Currently only tested to be working on Windows 10 x64 (64-bit) and Microsoft Outlook v16.0.

Before everything else, please ensure that you have logged in into your Microsoft Outlook account on your computer's Outlook desktop application (not the web client nor the Mail desktop program) and that you currently have and are connected to a working Internet connection.

Folders and rules are identified by name. As such, later "addition" updates of this software are automatically compatible to previous versions (except if there are name-related changes). However, it might not necessarily be compatible with other user-defined custom rules. As such, please also ensure that you currently do not have any "conflicting" rules in your Microsoft Outlook profile. A clean slate would be the best recommended state, but if that is not possible/desirable, then you can check against the [catalogue section](#catalogue) below to ensure that there are no conflicting rules.

You can download the ZIP file (latest version) from the [Releases](https://github.com/jamestiotio/SUTDigest/releases/latest) page and then run (double-click) the binary executable file (`SUTDigest.exe`). Wait for a while as the program executes the corresponding instructions. You can know that the program is done once the folders and the rules are present in your Outlook and a system indicator 'beep' sound on Windows can be heard.

The program will attempt to search and open your Microsoft Outlook desktop application, so do not worry if you see your Outlook window being opened/closed.

Alternatively, you can clone this repository, build the executable yourself and run it.

## Catalogue

### Folders

- Others
  - Student Clubs
  - Office of Marketing
  - Career Development Centre
  - Whatz Newz
  - Library
  - Student Government
  - All Students
  - News Coverage
  - Social Media Report
  - Office of Research
  - Office of International Relations
  - HASS Events
  - LKYCIC
  - AI Research
  - SMT
  - Wellbeing Services
  - UROP
  - UTOP
  - IAP Admin

### Rules

> To clarify, server-side rules are actually set up, stored, and deployed on the Outlook server (and hence would still operate on multiple different mail client alternatives like the Mail app or on the web and even after the local Outlook mail client is closed) since such rules would be executed on the corresponding Exchange server, while client-only rules only work when the local Outlook mail client is open.

- Server-Side:
  - Filter and move all emails with sender addresses that contain `club.sutd.edu.sg` to the Student Clubs folder.
  - Filter and move all emails with sender addresses that contain `marcom.sutd.edu.sg` to the Office of Marketing folder.
  - Filter and move all emails from [`careers@sutd.edu.sg`](mailto:careers@sutd.edu.sg) to the Career Development Centre folder.
  - Filter and move all emails with subjects that contain `Whatz Newz by Office of Student Life` to the Whatz Newz folder.
  - Filter and move all emails from [`library@sutd.edu.sg`](mailto:library@sutd.edu.sg) to the Library folder.
  - Filter and move all emails with sender addresses that contain `studentgov.sutd.edu.sg` to the Student Government folder.
  - Filter and move all emails sent to [`*AllStudents@sutd.edu.sg`](mailto:*AllStudents@sutd.edu.sg) to the All Students folder.
  - Filter and move all emails with subjects that contain `Summary of coverage: Highlights of SUTD` to the News Coverage folder.
  - Filter and move all emails with subjects that contain `SUTD's Social Media Report` to the Social Media Report folder.
  - Filter and move all emails from [`research@sutd.edu.sg`](mailto:research@sutd.edu.sg) to the Office of Research folder.
  - Filter and move all emails from [`global@sutd.edu.sg`](mailto:global@sutd.edu.sg) to the Office of International Relations folder.
  - Filter and move all emails from [`hassevents@sutd.edu.sg`](mailto:hassevents@sutd.edu.sg) to the HASS Events folder.
  - Filter and move all emails from [`lkycic@sutd.edu.sg`](mailto:lkycic@sutd.edu.sg) to the LKYCIC folder.
  - Filter and move all emails from [`AI_RESEARCH@sutd.edu.sg`](mailto:AI_RESEARCH@sutd.edu.sg) to the AI Research folder.
  - Filter and move all emails from [`sci-math@sutd.edu.sg`](mailto:sci-math@sutd.edu.sg) to the SMT folder.
  - Filter and move all emails from [`wellbeing@sutd.edu.sg`](mailto:wellbeing@sutd.edu.sg) to the Wellbeing Services folder.
  - Filter and move all emails from [`urop@sutd.edu.sg`](mailto:urop@sutd.edu.sg) to the UROP folder.
  - Filter and move all emails from [`utop@sutd.edu.sg`](mailto:utop@sutd.edu.sg) to the UTOP folder.
  - Filter and move all emails from [`iapadmin@sutd.edu.sg`](mailto:iapadmin@sutd.edu.sg) to the IAP Admin folder.

## Feedback

If you encounter any problems or if you would like to propose any new rules/folders, feel free to raise an issue or create a pull request! If you see any feature that you would like to be implemented, feel free to browse the issue/pull request list and vote for it by using emojis!

Feel free to also contribute by testing and verifying that this program works on other flavors/environment configurations of Windows. I doubt that it will work on other OS-es since it makes use of Windows-specific APIs. A platform-agnostic solution would need to somehow utilize the Microsoft Graph API, which is not accessible at the moment. 😔 (If you have any suggestions on making this OS-agnostic without Microsoft Graph API, feel free to bring it up!)

Please be aware that any new rules/folders should benefit the majority of the population and this is up to the discretion of the community. Rules/folders that are more personal in nature would not be accepted and should be individually applied/managed anyway.

## Disclaimer

I am not responsible for any breakage that this program does to your Microsoft Outlook account, Windows-based machine, etc. Do not blindly trust any code from random strangers and always verify them, if you have the technical knowledge to do so.

That said, obligatory formal note here:

The information, software, products, and services included in or available through this SUTDigest repository may include inaccuracies or typographical errors. Changes are periodically made to this repository and to the information therein. The creator/maintainer and/or the respective contributors may make improvements and/or changes in this repository at any time. Advice received via this repository page should not be relied upon for personal, medical, legal or financial decisions and you should consult an appropriate professional for specific advice tailored to your situation.

## Future TO-DOs

- Investigate about the best way to retroactively execute all rules on all messages in the current user's Inbox folder by calling the [`Execute()` function](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._rule.execute).
- Add more rules.

<p align="center">&mdash;⭐&mdash;</p>
<p align="center"><i>Crafted, designed and built with ❤️ by <a href="https://github.com/jamestiotio">@jamestiotio</a> in Singapore. (*^▽^*)</i></p>
