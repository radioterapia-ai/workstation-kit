# Third-party components

**None.**

The application is a single PowerShell 5.1 file and uses only what ships with
Windows:

| | |
|---|---|
| `System.Windows.Forms`, `System.Drawing` | the window, from the .NET Framework that ships with Windows |
| CIM / WMI | process, service and disk enumeration |
| PowerShell cmdlets | `Get-Process`, `Get-Service`, `Get-ScheduledTask`, `Get-CimInstance`, `Get-ItemProperty` |

There is no bundled library, no package manifest, no download at run time and no
self-update.

That absence is deliberate. A script that fetches a component at run time is
behaviour corporate antivirus software treats as suspicious, correctly. It also
means there is no dependency to audit, no supply chain to compromise, and nothing
to pin.

## What is not ours

The **names** in the protection lists — treatment planning systems,
record-and-verify systems, image viewers, electronic health records, remote access
tools, backup agents, security agents. They belong to their vendors and appear
here for one purpose: so that those programs are **not closed**. No code from any
of them is included, and no authorship over them is claimed.

## The fonts in the banner

`docs/img/banner.png` is a rendered image. The typefaces used to render it —
Outfit and JetBrains Mono — are licensed under the SIL Open Font License and the
Apache License 2.0 respectively, by their authors. The image is a work product;
no font file is redistributed in this repository.
