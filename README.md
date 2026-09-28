# Workstation Kit

> **Support tool. Not validated for clinical use.**
> This is not a medical device. It does not diagnose patients, does not measure,
> does not interpret medical images and does not replace professional judgement.
> It is a Windows utility. See [NOTICE](NOTICE).

Diagnoses, cleans and optimises the **open Windows session**. A single PowerShell
file, no installer, and **no administrator privilege required**.

Part of the **[Radioterapia.AI](https://radioterapia.ai)** ecosystem.

<p align="center">
  <img src="docs/img/banner.png" alt="Workstation Kit — diagnose, clean and optimise the open Windows session" width="100%">
</p>

---

## What it is, and when to run it

It is a **Windows app**, not a radiotherapy app. It prepares the machine so the
day's work can start — which is why it is meant to run **right after you sign in**,
before opening whatever you came to the computer to use.

Running it **in the middle of the day** can close work in progress. The app
measures that and says so: how long the session has been open, whether a batch is
running, and whether an application with an open window is already in use.

> The protection lists are the **net**, for anyone who runs it at the wrong time.
> The main mechanism is telling you the moment is wrong, with the reason in front
> of you, and letting you decide.

<p align="center">
  <img src="docs/img/tela.png" alt="The application window: module buttons on the left, coloured log on the right" width="100%">
</p>

## Nothing is permanent, except the files you choose to delete

| what | comes back through |
|---|---|
| closed processes, CPU priority, compacted memory | the next sign-in |
| stopped scheduled tasks | their own next scheduled run |
| items removed from startup, performance tweaks | the **Undo** button |
| old Downloads | the Recycle Bin |
| temporary files and caches | they do not come back — that is the cleanup |

The Recycle Bin is emptied **first**, and what leaves Downloads goes into it
afterwards, which is why it remains recoverable.

**It does not uninstall anything.** When a program looks dispensable, the app
**ends the process** instead of uninstalling it: the memory comes back now, and at
the next sign-in everything is in place again. Detection that would require a
permanent action — a dead network mapping, an orphaned credential — comes out as a
**finding with the command**, for you to decide.

Every action group has a declared way back, and the test suite reads the
execution list from the source and fails if a new group appears without one.

## The modules

| | |
|---|---|
| **Diagnostics and inventory** | a portrait of the machine, and what can be resolved |
| **Large files and folders** | read-only; lists and gives a verdict |
| **Cleanup** | temporary files, caches, old Downloads |
| **Optimise the session** | closes what is not in use, adjusts priority, compacts memory |
| **Persistence** | what comes back on its own at every sign-in |

The log is the product: eight levels, each with its own colour, timestamps, and
the reason written next to every decision. What you see on screen is what you can
copy into a support ticket.

## What it will not do

- does not require administrator, in any function
- does not uninstall anything, does not touch `HKLM`
- does not delete files outside cache folders in your own profile — the large-file
  module lists and gives a verdict; **you** delete, through Explorer
- does not clean the registry: there is no published evidence of measurable gain,
  and there is real risk
- does not download anything and does not update itself
- does not work around antivirus software

## Protecting work in progress

The app carries lists of processes it will never close. They exist because this
tool was born inside a radiotherapy service, where closing the wrong process does
not mean an inconvenience — it means interrupting a treatment session.

Those lists cover treatment planning systems, record-and-verify systems, image
viewers, electronic health records, published-application clients, remote-access
tools used by IT, backup agents and corporate security agents, from several
vendors. They are the only part of this app that touches the medical world, and
they cost nothing to anyone who does not need them.

> The direction of the error is not symmetric. Protecting something that did not
> need it leaves dead weight. Failing to protect something that did interrupts
> someone's work. Doubt is resolved by protecting.

## Limitations

Stated here because a protection list that claims to be complete is worse than one
that says where it stops.

**The lists are incomplete, and always will be.** Process names for clinical
software barely exist in public sources: vendors publish brochures and conformance
statements, not executable names. The tokens in the lists are **product names, not
measured process names** — a token protects only if the real executable name
contains it. Coverage here means *declared*, not *verified on a real installation*.

A longer list is not the answer, and three stronger mechanisms are **not
implemented** yet: matching by vendor folder, by executable publisher
(`CompanyName` or Authenticode signature), and by licence manager.

**Categories nothing covers today**, all of which run in the user's session:
medical dictation, smartcard and digital-signature middleware, **accessibility
tools** (a screen reader has no main window, so it arrives pre-selected — closing
it leaves a blind user with no interface), third-party input method editors, and
VPN clients beyond the four listed.

Two cases worth naming:

- **Local neural-network auto-contouring** uses the workstation CPU for one to
  three minutes, which to an optimiser looks exactly like a process hogging the
  machine. It is the first candidate to be closed by mistake, and what it
  interrupts is a contour in progress.
- **ClickOnce applications** install into `%LOCALAPPDATA%\Apps\2.0`. A cleaner
  that treats that tree as cache can delete the installation.

**Scope.** No administrator, so session 0 is out of reach — services are protected
by scope rather than by list. It uninstalls nothing, cleans no registry, downloads
nothing and does not update itself.

## Languages

**English, Portuguese and Spanish.** Three buttons at the bottom of the left
panel switch the interface at once — no restart, and the log already written
stays in the language it was written in, because it is a record of what happened
rather than a screen to redraw.

The first run guesses from the Windows UI culture and falls back to English. The
choice is remembered in `HKCU`.

Portuguese is the **source**: those entries are the original text, and the other
two are translations of it. The build gate enforces **placeholder parity** — a
translation with fewer `{N}` than the source would make the format operator
throw, and this app runs with a hidden console where a thrown error ends the step
in silence.

Right-to-left languages are deliberately out of scope, for layout rather than
market reasons: this app's content is largely Windows paths, process names and
versions, which the bidirectional algorithm reorders inside a right-to-left
paragraph until a path is unreadable.

## Running it

Download the release, unzip both files into the same folder, and run:

```
Start_WorkstationKit.cmd
```

PowerShell 5.1, Windows 10 or 11. Nothing is installed, nothing is downloaded,
and no administrator prompt appears. To remove it, delete the two files.

Every release is built behind a gate that has to pass first: the protection
lists, the language layer's placeholder parity, the environment traps, and a
static defect hunt over the source. The packager then reopens the built archive
and checks that the file inside it parses, kept its byte-order mark, and matches
the source hash.

---

## The ecosystem

**[Radioterapia.AI](https://radioterapia.ai)** is built across four layers, and
they are deliberately separate — what runs on a clinical workstation has different
constraints from what runs in a browser.

| layer | what lives there |
|---|---|
| **Web** — [radioterapia.ai](https://radioterapia.ai) | an **AI-first hub for medical skills**, and a hub for applications and community contributions. Expert Mode for professionals, Patient Information mode in 12 languages |
| **Local** | what runs on the clinic's own machine, because the data must not leave it: this kit, auto-contouring, local pseudonymization |
| **Mobile** | Android in the room — [PhotoID RT](https://github.com/radioterapia-ai/photoid-rt) |
| **[Hugging Face](https://huggingface.co/Radioterapia-AI)** | what needs heavy dependencies a common user should not have to install, for quick use straight from the browser — [POP de Elite](https://huggingface.co/spaces/Radioterapia-AI/POP) and [Fábrica de Slides](https://huggingface.co/spaces/Radioterapia-AI/Fabrica_de_Slides) |

The layers are not tiers of the same thing, and one move shows why. **ContourLab
used to run on Hugging Face and does not any more** — the space that carries its
name is now a signpost pointing to the Local Suite. Auto-contouring reads patient
images, and patient images belong on the clinic's machine. The capability did not
improve when it moved; the place did.

That is the rule the whole ecosystem is arranged around: **the layer is chosen by
what the data is, not by what is convenient.** A tool that never sees patient data
is better off in a browser, where nobody has to install anything. A tool that does,
runs where the data already lives.

The web layer is where the community comes in: list your app, share your
repository, or deploy with us. See **[radioterapia.ai/about](https://radioterapia.ai/about)**.

## Who builds this

**Radioterapia.AI** — *Inteligência além das fronteiras da saúde.*

**Dr. Henrique Faria Braga** — Medical Skills & Apps. Radiation oncologist with
more than ten years building automation for operational and managerial workflows
in radiation therapy, and the originator of Radioterapia.AI. Medicine at FMUSP,
residencies in Internal Medicine and Radiation Oncology at USP. Head of the
Radiation Therapy team at Rede Américas and Medical Coordinator of Oncology at
Centro Médico Samaritano Barra da Tijuca, Rio de Janeiro.
CREMESP 129263 · CREMERJ 52-111804-8 · RQE-SP 54873 · RQE-RJ 331440 · CNEN CB-8319

**Fís. Lucas Brito** — Physics Scripts & Architecture, **co-founder**. Medical
physicist with a doctorate in Medical Physics, responsible for the platform's
technical architecture and AI tooling. He registered the domain and runs the
infrastructure that put the web ecosystem online.

### Links

| | |
|---|---|
| Website | [radioterapia.ai](https://radioterapia.ai) · [about](https://radioterapia.ai/about) |
| Instagram | [@radioterapia.ai](https://www.instagram.com/radioterapia.ai/) · [@radioterapiabr](https://www.instagram.com/radioterapiabr/) · [@podirradiar](https://www.instagram.com/podirradiar/) |
| LinkedIn | [Henrique Braga](https://www.linkedin.com/in/henriquefbraga/) · [Lucas Brito](https://www.linkedin.com/in/lucassbrito/) |
| Personal site | [drhenriquebraga.com.br](https://drhenriquebraga.com.br/) |

---

## Licence and name

The source code is released under the **[Apache License 2.0](LICENSE)**.

Two things travel with it and are not optional:

- **[NOTICE](NOTICE)** — section 4(d) of the licence requires it to accompany every
  redistribution. It carries the declaration that this is not validated for
  clinical use.
- **The name is not licensed.** Section 6 grants no trademark rights:
  "Radioterapia.AI" and "Workstation Kit", and their logos, may not be used to
  identify derived products.

Copyright © 2026 Henrique Faria Braga. Radioterapia.AI is a trade name and a
website, not a legal entity.

Learn more at **[radioterapia.ai](https://radioterapia.ai)**.
