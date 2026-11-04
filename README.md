# Ramarc Weight Scanner

Phone app for Ramarc Foods receiving: scan box and combo labels with the phone camera (or a Bluetooth scanner), check combos against the Ramarc scale, and send a weight-sheet PDF.

Open: https://jesusramarc-dot.github.io/ramarc-scanner/

Scans are kept on each phone only; nothing is uploaded to this site.


## Drafts and sharing between phones

- **Save draft & switch sheet** parks a weight sheet so you can work on several at once.
- **Weight sheets in progress** (start screen) lists drafts from every connected phone; **Completed weight sheets** can be reopened to fix and re-send.
- Sharing between phones uses a small Google helper in Ramarc's Google account: [`helper/Code.gs`](helper/Code.gs). Setup steps are at the top of that file. Every phone connects to it automatically.
