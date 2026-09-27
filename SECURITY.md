# Security policy

## Reporting a vulnerability

**Please do not open a public issue.**

Use GitHub's private vulnerability reporting on this repository — the **Security**
tab, then *Report a vulnerability*. It goes only to the maintainer and does not
require exchanging email addresses.

If that is unavailable to you, get in touch through
**[radioterapia.ai](https://radioterapia.ai)** and say only that you have a
security report; the details can then move to a private channel.

You will get a first reply within **7 days**.

## What is in scope

This tool closes processes, stops scheduled tasks and deletes cache files on a
workstation that may be in clinical use. The findings that matter most are the
ones that make it touch something it should not:

- Anything that makes a protected process closeable — a treatment planning
  system, a record-and-verify system, an electronic health record, a published
  application client, a remote-access tool, a backup agent, a security agent.
- Anything that makes a cleanup target resolve outside the caller's own profile,
  or that makes an empty configuration value expand into a wildcard that matches
  every path.
- Anything that makes the app act without the user having seen and selected the
  item first.
- Anything that sends data off the machine. The app has no network output by
  design: it does not download, does not update itself, and writes only to the
  inventory location the user configures.
- Anything that lets a value read from the environment — a file, a registry key,
  a process name — change what the app executes.

## What is already known, and is not a finding

- **It is not a medical device and is not validated for clinical use.** A report
  that it produced a clinically wrong judgement is a correctness bug, not a
  vulnerability. Open an issue — it is welcome.
- **The protection lists are incomplete**, and [docs/LIMITATIONS.md](docs/LIMITATIONS.md)
  says where. A vendor that is not covered is a gap we already publish, not a
  disclosure.
- **It requires no administrator, and therefore cannot reach session 0.** System
  services are out of its reach by scope, not by list.
- **Antivirus software may flag it.** It is an unsigned PowerShell script that
  enumerates processes. The project does not obfuscate to avoid detection; the
  answer is for IT to allow it by hash after reading the source.

## Supported versions

The latest release. Single maintainer, no backports.

## The gate

Every push and pull request runs `.github/workflows/gate.yml` on Windows
PowerShell 5.1: the protection suite, the language layer, a leak scan for
third-party network identifiers, and the environment traps. A red gate blocks
the merge. You can run the same thing locally:

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Tudo.ps1
```
