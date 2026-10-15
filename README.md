# Thread Attachments for Outlook

An Outlook add-in that lists every attachment in an email thread, lets you preview them, and combines the ones you pick into a single PDF to save or print.

**Status:** test build 0.2.

## Install

Open this repository's GitHub Pages site (the address shown under **Settings → Pages**) and follow the steps there. The page builds a manifest file that points Outlook at that site.

## What it does

- Lists each attachment once, oldest first, across all the selected emails in a thread.
- Select all, or pick individual files.
- Previews PDFs, photos and multi-page fax scans (TIFF) in the side panel.
- Combines PDFs, photos and scans into one PDF, then saves it, prints it (choose *Save as PDF* to save a file), or puts it into a new email.
- Pick individual pages of a PDF in the preview, for example the POD but not the rate confirmation sent in the same file.

## Privacy

Everything runs inside Outlook's side panel. Attachments are read and combined on your computer and are never uploaded. This site only hosts the add-in's code.

Built with [pdf-lib](https://pdf-lib.js.org), [PDF.js](https://mozilla.github.io/pdf.js/) and [UTIF.js](https://github.com/photopea/UTIF.js). See `THIRD-PARTY-NOTICES.txt`.
