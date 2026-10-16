# Thread Attachments for Outlook

An Outlook add-in that lists every attachment in an email thread, lets you preview them, and combines the ones you pick into a single PDF to email or save.

**Status:** test build 0.11.

## Install

Open this repository's GitHub Pages site (the address shown under **Settings → Pages**). The team admin creates a team install file there once, publishes the team templates from inside Outlook, and sends the install file to reps.

## What it does

- Lists each attachment once, oldest first, across all the selected emails in a thread, with a small picture of each file (page 1 of a PDF or fax, or the photo) and type icons for the rest.
- Pick individual files, or Select all. Nothing is checked until you pick.
- Previews PDFs, photos and multi-page fax scans (TIFF) in the side panel.
- Combines PDFs, photos and scans into one PDF, then emails it or saves it under a name you type.
- Email templates: team templates managed by one admin. Reps can add their own only once the admin allows it. Every email asks for the load number, names the PDF after it, fills in the recipients and message, and uses the subject "<load> POD(s)".
- Per-template options: say "No Lumper" when Lumper isn't ticked; each stop in its own file or all stops in one; lumper receipts in their own file (tag them per file, or per page in the preview).
- Template checklists: items that have to be ticked before the email is created.
- Template questions: a pop-up asked when the template is picked, before the email step. Each answer goes to the next question, jumps ahead to a later one, continues to the email, or stops with a message. The template's customer name fills in {company}.
- Rate confirmation check: the admin sets a sample rate confirmation; pages like it (digital, scanned or faxed) are flagged and must be left out (or cleared) before emailing or saving.
- Multi-stop: one PDF per stop, named load-1, load-2…, whatever their size. Each checked file gets a stop number, and a PDF holding several stops can be split by picking a stop for each page in the preview.
- Accessorials: tick Detention (shipper or receiver, IN, APPT, OUT, hours requested), Lumper (total) or Other (type, total) when creating the email, and they're listed in the message. On a Multi-stop load, the receiver choices name the stop, like "Receiver (Stop 2)".
- When only one email comes through, a tip says so, and its info icon plays a short how-to for selecting every email in the thread (new Outlook or classic Outlook).
- Pick individual pages of a PDF in the preview, for example the POD but not the rate confirmation sent in the same file.
- Size limits: PDFs over a template's (or Save PDF's) limit are split into name-1, name-2… keeping page order.
- Optional automatic attaching: with the team admin's PDF Emails Azure function (its own Function App, kept outside this site), Create email opens with the PDF already attached.

## Privacy

Everything runs inside Outlook's side panel. Attachments are read and combined on your computer. Only when the team admin turns on automatic attaching are PDFs made with Email uploaded, to the admin's PDF Emails Azure function, for Outlook to download into the new email. Download links stop working after an hour, and expired files are deleted. `team.json` holds the team templates and settings encrypted with the team key, which only the team install file contains. Personal templates are saved in each person's Outlook account.

Built with [pdf-lib](https://pdf-lib.js.org), [PDF.js](https://mozilla.github.io/pdf.js/) and [UTIF.js](https://github.com/photopea/UTIF.js). See `THIRD-PARTY-NOTICES.txt`.
