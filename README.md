# Restack for Outlook

An Outlook add-in that lists every attachment in an email thread, lets you preview them, and combines the ones you pick into a single PDF to email or save. Restack was called Thread Attachments before 0.13; it's the same add-in (same ID), so team keys and published templates carry over.

**Status:** test build 0.16.

## Install

Open this repository's GitHub Pages site (the address shown under **Settings → Pages**). The team admin creates a team install file there once, publishes the team templates from inside Outlook, and sends the install file to reps. Teams that already use Thread Attachments get the new install file from the add-in instead (**Settings → Team → Share install file**), which keeps the team key.

To try a new version before reps get it, put it in the repository's `beta` folder and use **Restack Beta** (**Settings → Team → Share beta install file**), which installs next to Restack and runs whatever is in that folder.

## What it does

- Lists each attachment once, oldest first, across all the selected emails in a thread, with a small picture of each file (page 1 of a PDF or fax, or the photo) and type icons for the rest.
- Pick individual files, or Select all. Nothing is checked until you pick.
- Set the order files go into the PDF: checked files are numbered; drag one by its number or use its up/down arrows. Files tagged Receipt (or tagged pages) always go after the PODs, or into their own file for templates that want that.
- Previews PDFs (including fax-style and JPEG 2000 scans), photos and multi-page fax scans (TIFF) in the side panel. Move a scan's pages into the right order with the arrows on each page; the attachment itself isn't changed.
- Photos go into the PDF at most 2000 pixels on their long side: the whole photo, smaller, never cropped.
- Combines PDFs, photos and scans into one PDF, then emails it or saves it under a name you type.
- Every PDF stays under 10 MB: a bigger one is split into name-1, name-2… keeping page order, all attached to the same email. Stop and receipt files are kept whole, with a note if one is over.
- Email templates: team templates managed by one admin. Reps can add their own only once the admin allows it. Every email asks for the load number, names the PDF after it, fills in the recipients and message, and uses the subject "<load> POD(s)".
- Per-template options: say "No Lumper" when Lumper isn't ticked; each stop in its own file or all stops in one; receipts in their own file instead of after the PODs (tag them per file, or per page in the preview).
- Accessorials in a separate email (per template): when Detention, Layover, Lumper or Other is ticked, Create email opens two emails. The POD email has the usual subject and the PODs. The request ("<load> - Detention/Lumper Request") has the PODs and receipts, the accessorials' details, and its own To, CC and message (a blank To uses the POD email's; a blank CC means no CC).
- Template questions: a pop-up asked when the template is picked, before the email step. Each answer goes to the next question, jumps ahead to a later one, continues to the email, or stops with a message. {company} fills in with the template's name.
- Rate confirmation check: the admin sets a sample rate confirmation (up to 3 pages). Pages like it (digital, scanned or faxed, compared by words and by layout), pages that say Rate Confirmation or Load Confirmation, and pages with the same barcode number as a flagged page are flagged and must be left out (or cleared) before emailing or saving.
- Multi-stop: one PDF per stop, named load-1, load-2…. Each checked file gets a stop number, and a PDF holding several stops can be split by picking a stop for each page in the preview.
- Accessorials: tick Detention (shipper or receiver, IN, APPT, OUT, hours requested), Layover (shipper or receiver, IN, APPT, OUT), Lumper (total) or Other (type, total) when creating the email, and they're listed in the message. On a Multi-stop load, the receiver choices name the stop, like "Receiver (Stop 2)".
- When only one email comes through, a tip says so, and its info icon plays a short how-to for selecting every email in the thread (new Outlook, classic Outlook, or mail not shown as conversations).
- Keep adding (for mail not shown as conversations): with the panel pinned, each email clicked adds its files to the list. Restack asks first when one looks like a different load.
- Pick individual pages of a PDF in the preview, for example the POD but not the rate confirmation sent in the same file.
- Rotate pages and photos in the preview, a quarter turn per click, so sideways scans and photos go into the PDF upright. The attachment itself isn't changed.
- Settings: the gear at the bottom of the panel, there whatever email is selected (even one without attachments, or none). It has Manage templates (for whoever can make templates), Team (the admin's: publish, install files, reps' own templates, rate con check, automatic attaching, receipt file names, usage counts and the admin console), Appearance (Light, Dark or Match Outlook), Text size (Small, Medium, Large), View (Large, Medium or Small pictures, List, Details, Tiles), and switches for Tag receipts by file name and Keep adding. Each person's choices are saved in their Outlook account.
- Templates are listed A to Z. Pin the ones you use most to the top of the Email menu and put them in your own order; the rest are under Other templates. A new template starts with every box unticked, and receipts in their own file are named "{load} Receipt".
- Tag receipts by file name: files with the admin's words in their name (lumper and receipt to start) start out tagged Receipt.
- Admin console (console.html, opened from Settings → Team): the team's templates and settings at full size in a browser, publishing straight to GitHub with an optional token, and a Usage tab.
- Usage counts: with the admin's Restack Usage Azure function (its own Function App, kept outside this site), each panel counts the PDFs saved, emails created, files, pages and rate cons caught, shown on the console's Usage tab by day, rep and template.
- Optional automatic attaching: with the team admin's PDF Emails Azure function (its own Function App, kept outside this site), Create email opens with the PDF already attached.

## Changing the wording

Most labels reps see on the panel (Select all, Receipt, Email, Create email, Detention, Include in PDF and so on) are in one list at the top of `taskpane.js`, called `WORDS`. Change the text between the double quotes, keep anything in {curly brackets}, and upload the file. An apostrophe is fine; a double quote inside the text isn't. Each label changes everywhere it's used, including the email's accessorial lines, within a few minutes (or close and reopen Outlook). Messages and warnings, and the template screens, aren't in the list.

## Privacy

Everything runs inside Outlook's side panel. Attachments are read and combined on your computer. Only when the team admin turns on automatic attaching are PDFs made with Email uploaded, to the admin's PDF Emails Azure function, for Outlook to download into the new email. Download links stop working after an hour, and expired files are deleted. When the team admin turns on usage counts, each panel sends counts to the admin's Restack Usage Azure function: the rep's email address, when, how many files, pages and PDFs, the template's name and which accessorials. Never files, file names, load numbers or customer addresses. `team.json` holds the team templates and settings encrypted with the team key, which only the team install file contains. Personal templates are saved in each person's Outlook account.

Built with [pdf-lib](https://pdf-lib.js.org), [PDF.js](https://mozilla.github.io/pdf.js/) and [UTIF.js](https://github.com/photopea/UTIF.js), set in [Archivo](https://github.com/Omnibus-Type/Archivo). See `THIRD-PARTY-NOTICES.txt`.
