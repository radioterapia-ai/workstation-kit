# Contributing

A fork that goes its own way is a fine outcome. So is an issue that only reports
what happened on one machine — the field report is how most defects here were
found, not code review.

## The gate

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Tudo.ps1
```

Four checks, and any non-zero exit blocks the merge. The same thing runs in
`.github/workflows/gate.yml` on every push and pull request.

| check | what it proves |
|---|---|
| `Testar-Sessao.ps1` | parse, no orphan calls, the eight protection lists load intact, and the negative cases still fail where they must |
| `Testar-Idioma.ps1` | every key exists in every language, no entry is empty, and **placeholder parity** holds |
| `Verificar-Vazamento.ps1` | no third-party network identifier reaches a public repository |
| `Testar-Ambiente.ps1` | BOM, line endings, no `chcp` in a `.cmd`, no hard-coded drive letter |

## Four things that will get a change rejected

**1. It requires administrator.** The scope of this tool is what a user can do
without elevation. A feature that only works elevated becomes a finding in the
diagnostics, with the text the user takes to their IT desk.

**2. It is not reversible.** Nothing may survive a restart except files the user
chose to delete. Block 16 of the suite reads the execution list from the source
and fails on any group without a declared way back.

**3. It splits the `.ps1` into modules.** The distribution is a file copy, not an
installation. One file is what makes that work, and it is decided.

**4. It adds a protection-list entry without a test case first.** Write the case,
watch it fail, then add the entry. It is the case that exposes the entry that does
not match — wrong spelling, missing anchor, a name too short to be safe unanchored.

## Language

| layer | language |
|---|---|
| what the user reads on screen | English, Portuguese, Spanish |
| code, comments, commit messages | Portuguese |
| README, this file, `SECURITY.md`, `docs/` | English |

Portuguese is the **source** for the interface strings: `pt` entries are the
original text, and the other two are translations of it. Never edit an `en` or
`es` entry to say something the `pt` entry does not.

## Adding a language

1. Add `[pscustomobject]@{ Cod = 'xx'; Nome = '...' }` to `$script:Idiomas`.
2. Copy the `pt` table to `$script:Textos['xx']` and translate the values.
3. Run `tools\Testar-Idioma.ps1`. It will tell you exactly which keys are empty
   and which have a placeholder set that does not match the source.

**Placeholder parity is the one that matters.** A translation with fewer `{N}`
than the source makes the format operator throw, and this app runs with a hidden
console: the step would end with no message at all.

## Reporting a problem

Include the Windows version and UI language, the interface language you had
selected, and the log — the *Copy log* button puts it on the clipboard. If it
involves a process that should not have been closed, include its name exactly as
Task Manager spells it, and the folder it runs from.

Security issues: see [SECURITY.md](SECURITY.md) and use private vulnerability
reporting rather than a public issue.

## Commits

Say what changed and why it was wrong before. A commit message that only names
the file has thrown away the reason, and the reason is the part that stops the
defect coming back.
