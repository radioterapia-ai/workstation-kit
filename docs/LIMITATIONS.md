# Limitations

What this tool does not cover. Published because a protection list that claims to
be complete is worse than one that says where it stops.

---

## The protection lists are incomplete, and always will be

**Process names for radiation therapy software barely exist in public sources.**
Vendors publish brochures, specifications and DICOM conformance statements; they
do not publish executable names. What is covered came from uninstaller pages and
executable catalogues, never from the vendor.

Measured as closeable today, among others: Pinnacle3, TomoTherapy, Accuray,
Brainlab Elements, SNC Patient, VeriSoft, myQA, Delta4, MobiusFX, ClearCheck,
Limbus, RadiAnt, Orthanc, dcm4chee.

> An incomplete list is the serious failure mode here, so the answer is not a
> longer list. It is four independent mechanisms, three of which do not require
> knowing the vendor.

| | mechanism | status |
|---|---|---|
| **A** | process name, anchored where short | implemented |
| **B** | **vendor folder** — a process whose path falls under `\Varian\`, `\CMS\`, `\Elekta\`, `\RaySearch\`, `\Accuray\`, `\Brainlab\` | **not implemented** |
| **C** | **executable publisher** — `CompanyName` or the Authenticode signature. Reads a file; needs no administrator | **not implemented** |
| **D** | **licence manager** — `hasplms`, `CodeMeter`, `lmgrd`, `FNPLicensingService`. Killing one does not bring the application down immediately; it brings it down at the next checkout, with a message about licensing and never about a closed process | **not implemented** |

## Six categories nothing covers, all closeable today

All of them run **in the user's session**, which is exactly where this tool acts:

1. **Medical dictation** — `natspeak`, `DragonBar`. Today `nssystem` survives *by
   accident*, matching the broad token `System`.
2. **Digital signature and smartcard middleware** — `SafeSignIC`,
   `SafeNetAuthentication`. Without it, a plan approval cannot be signed.
3. **Accessibility** — `nvda`, `jfw`, `Narrator`. A screen reader runs with no
   main window, so it arrives **pre-selected**. Closing it leaves a blind user
   with no interface.
4. **Third-party IME** — `SogouInput`, `Baidu`, Google Japanese Input. In an app
   with a language selector, closing the IME is an app that does not work in
   several of them.
5. **Backup agents.** Proof that current coverage is accidental: `AcronisAgent`
   used to survive only because it matched the token `Onis`, from an unrelated
   DICOM viewer. Designed entries now exist; the accidental match remains as
   redundancy.
6. **VPN beyond the four listed** — `openvpn`, `wireguard`, `NetExtender`.
   Dropping one cuts access to the image archive **and the very session the tool
   is probably being run through**.

## Two cases worth naming

- **Neural-network auto-contouring that runs locally** (Limbus Contour,
  Radformation AutoContour) uses the workstation CPU for one to three minutes. To
  an optimiser that has exactly the signature of "a process hogging the machine".
  It is the first candidate to be closed by mistake, and what it interrupts is a
  contour in progress.
- **ClickOnce applications** (Sectra IDS7 among them) install into
  `%LOCALAPPDATA%\Apps\2.0`. A cleaner that treats that tree as cache can delete
  the installation.

## The language layer

Three languages: English, Portuguese, Spanish. Portuguese is the source; the
other two are translations of it.

**Right-to-left languages are deliberately out of scope** — not for market
reasons, for layout. This app's content is almost entirely Latin and
untranslatable: Windows paths, process names, versions. The bidirectional
algorithm visually reorders those runs inside a right-to-left paragraph, and a
path with backslashes and digits becomes unreadable. This is an application about
paths, where the user reads the path and decides whether to delete.

## What the leak scanner does not catch

Declared in its own header, repeated here because it bounds a promise:

- a corporate name **without** a TLD — it has the shape of any word
- a UNC to a **short** host without digits — requiring less than that fired on
  regular expressions with escaped backslashes
- anything inside a binary file
- a person's name, which has no shape at all

## Scope, stated plainly

- No administrator, so **session 0 is out of reach**. Services are protected by
  scope, not by list — including most database engines and image servers.
- **It uninstalls nothing.** Where a cleaner would tell you to uninstall, this
  closes the process instead: the memory comes back now and the program is back
  at the next sign-in.
- It **does not clean the registry**. No published evidence of measurable gain,
  and real risk.
- It **downloads nothing and does not update itself.**
- Not a medical device. Not validated for clinical use.
