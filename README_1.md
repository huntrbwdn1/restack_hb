# Restack for Outlook

An Outlook add-in that lists every attachment in an email thread, lets you preview them, and combines the ones you pick into a single PDF to email or save. Restack was called Thread Attachments before 0.13; it's the same add-in (same ID), so team keys and published templates carry over.

**Status:** test build 0.14.

## Install

Open this repository's GitHub Pages site (the address shown under **Settings → Pages**). The team admin creates a team install file there once, publishes the team templates from inside Outlook, and sends the install file to reps. Teams that already use Thread Attachments get the new install file from the add-in instead (**Email ▾ → Manage templates → Share install file**), which keeps the team key.

## What it does

- Lists each attachment once, oldest first, across all the selected emails in a thread, with a small picture of each file (page 1 of a PDF or fax, or the photo) and type icons for the rest.
- Pick individual files, or Select all. Nothing is checked until you pick.
- Set the order files go into the PDF: checked files are numbered; drag one by its number or use its up/down arrows. Files tagged Receipt (or tagged pages) always go after the PODs, or into their own file for templates that want that.
- Previews PDFs (including fax-style and JPEG 2000 scans), photos and multi-page fax scans (TIFF) in the side panel.
- Combines PDFs, photos and scans into one PDF, then emails it or saves it under a name you type.
- Every PDF stays under 10 MB: a bigger one is split into name-1, name-2… keeping page order, all attached to the same email. Stop and receipt files are kept whole, with a note if one is over.
- Email templates: team templates managed by one admin. Reps can add their own only once the admin allows it. Every email asks for the load number, names the PDF after it, fills in the recipients and message, and uses the subject "<load> POD(s)".
- Per-template options: say "No Lumper" when Lumper isn't ticked; each stop in its own file or all stops in one; receipts in their own file instead of after the PODs (tag them per file, or per page in the preview).
- Accessorials in a separate email (per template): when Detention, Lumper or Other is ticked, Create email opens two emails. The POD email has the usual subject and the PODs. The request ("<load> - Detention/Lumper Request") has the PODs and receipts, the accessorials' details, and its own To, CC and message (a blank To uses the POD email's; a blank CC means no CC).
- Template checklists: items that have to be ticked before the email is created.
- Template questions: a pop-up asked when the template is picked, before the email step. Each answer goes to the next question, jumps ahead to a later one, continues to the email, or stops with a message. The template's customer name fills in {company}.
- Rate confirmation check: the admin sets a sample rate confirmation (up to 3 pages). Pages like it (digital, scanned or faxed, compared by words and by layout), pages that say Rate Confirmation or Load Confirmation, and pages with the same barcode number as a flagged page are flagged and must be left out (or cleared) before emailing or saving.
- Multi-stop: one PDF per stop, named load-1, load-2…. Each checked file gets a stop number, and a PDF holding several stops can be split by picking a stop for each page in the preview.
- Accessorials: tick Detention (shipper or receiver, IN, APPT, OUT, hours requested), Lumper (total) or Other (type, total) when creating the email, and they're listed in the message. On a Multi-stop load, the receiver choices name the stop, like "Receiver (Stop 2)".
- When only one email comes through, a tip says so, and its info icon plays a short how-to for selecting every email in the thread (new Outlook or classic Outlook).
- Pick individual pages of a PDF in the preview, for example the POD but not the rate confirmation sent in the same file.
- Settings: the gear at the bottom of the panel, there whatever email is selected (even one without attachments, or none). It has Manage templates, for whoever can make templates, and Appearance: Light, Dark or Match Outlook (the default), saved in each person's Outlook account.
- Templates are listed A to Z, in the Email menu and in Manage templates.
- Optional automatic attaching: with the team admin's PDF Emails Azure function (its own Function App, kept outside this site), Create email opens with the PDF already attached.

## Changing the wording

Most labels reps see on the panel (Select all, Receipt, Email, Create email, Detention, Include in PDF and so on) are in one list at the top of `taskpane.js`, called `WORDS`. Change the text between the double quotes, keep anything in {curly brackets}, and upload the file. An apostrophe is fine; a double quote inside the text isn't. Each label changes everywhere it's used, including the email's accessorial lines, within a few minutes (or close and reopen Outlook). Messages and warnings, and the template screens, aren't in the list.

## Privacy

Everything runs inside Outlook's side panel. Attachments are read and combined on your computer. Only when the team admin turns on automatic attaching are PDFs made with Email uploaded, to the admin's PDF Emails Azure function, for Outlook to download into the new email. Download links stop working after an hour, and expired files are deleted. `team.json` holds the team templates and settings encrypted with the team key, which only the team install file contains. Personal templates are saved in each person's Outlook account.

Built with [pdf-lib](https://pdf-lib.js.org), [PDF.js](https://mozilla.github.io/pdf.js/) and [UTIF.js](https://github.com/photopea/UTIF.js), set in [Archivo](https://github.com/Omnibus-Type/Archivo). See `THIRD-PARTY-NOTICES.txt`.
