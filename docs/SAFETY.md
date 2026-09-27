# Safety rules

What must not break. Every rule here came from a defect that shipped, not from a
precaution someone imagined.

---

## 0. The rule the others come from

> **Distrust every success message that does not check its own result.**

For weeks the tool printed *"Shortcuts copied to the Desktop!"* on every machine
without having copied anything. The path was wrong and the `xcopy` error was
silenced. It surfaced only when somebody added an honest counter.

## 1. Nothing requires administrator

Not a preference — the scope. A feature that only works elevated does not go in;
it becomes a finding with the text the user takes to their IT desk.

## 2. Run it right after signing in

This is a Windows utility, not a clinical one. It prepares the machine so the
day's work can start, which is why the moment matters: run in the middle of the
day, it can close work in progress.

> The protection lists are the **net**, for anyone who runs it at the wrong time.
> The main mechanism is `Get-MomentoSessao`, which measures and says so.

Three signals, none of them deciding alone: the age of **this** session (the
session's `explorer`, not machine uptime, which cannot tell "just signed in" from
"here since 7am"); ecosystem load running; and a work application open **with a
window**.

A window is required deliberately. Agents and tray icons start at sign-in without
anyone asking; counting them would make the warning fire on every run, and a
warning that always fires is a warning nobody reads.

## 3. Nothing is permanent

> **Nothing this app does survives a restart, except the files the user chose to
> delete.**

Every action group has a declared way back: the next sign-in, the Undo button
(which stores the previous value before writing), or the Recycle Bin. Two groups
were removed for failing this — `net use /delete` and `cmdkey /delete` do not come
back at all. They are now findings that print the command.

Block 16 of the suite reads the execution list from the source and requires every
group to appear in a classification **declared in the test**. A new group without
a way back fails. The classification is deliberately not read from the code: if it
were, the test would agree with whatever the code said.

## 4. The asymmetry that governs every list decision

> **Protecting something that did not need it leaves dead weight. Failing to
> protect something that did interrupts someone's work. The two errors do not
> cost the same, and doubt is resolved by protecting.**

One careless edit once marked a database engine, an image server and the remote
access tool as dispensable.

### Names in a regular expression match too much

The lists are regular expressions matched against `ProcessName`, unanchored by
default. Measured in this code: `ARIA` (an oncology information system) also
matches `Soarian` (an unrelated health record). `Epic` — meant for a game
launcher — matched Epic Systems, the largest electronic health record in the
world, and a match there meant the record came up **pre-selected for closing**.

> A three or four letter name in a protection list almost always needs `^` and `$`.

**Write the test case before adding the entry.** It is the case that exposes the
entry that does not match — wrong spelling, missing anchor, too short.

## 5. Stored data never decides to close a process

Module 5 queries the system once per operation instead of once per process. That
is what makes it fast, and it is where six defects came from at once.

> **A snapshot describes the instant it was taken. The decision to close happens
> later — and what changed in between is exactly what needs protecting.**

1. **Whoever decides reads live.** The path comes from `Get-CaminhoAoVivo`, which
   reads `.Path` at that moment and uses the snapshot only as backup, for a
   process whose `.Path` will not open. Never the other way round.
2. **Whoever stores declares an expiry.** 15 seconds. The read guard checks the
   **contents**, not just existence: a snapshot with no live process is not an
   answer, and neither side trusts the other to have behaved.
3. **A failed query is not stored.** Caching the empty result would disable the
   protection for the whole operation, silently. Note the language detail: an
   empty `pscustomobject` is truthy in PowerShell, so `if ($cache)` does not tell
   "I have an answer" from "I have an empty object".

## 6. Text the system returns is localised

One version filtered `sc query` on the word `STATE`. Windows in Portuguese returns
`ESTADO`. Result: **zero security agents** reported on every pt-BR station, with
no error at all.

> **A cmdlet that returns an OBJECT is safe. A command that returns TEXT is
> localised, and comparing that text is a defect that only appears on somebody
> else's machine.**

### And case folding is localised too

Measured in PowerShell 5.1: under `tr-TR` and `az-Latn-AZ`, `'CITRIX' -match
'citrix'` returns **False**. Turkish and Azerbaijani lowercase `I` to a dotless
`ı`, which is not the `i` in the pattern. The failure is **fail-open**: the line
is `if ($name -match $protected) { continue }`, so a failed match means it does
**not** skip — it closes.

This does not depend on the app offering Turkish. A Turkish workstation breaks the
app as it is. All decision points go through `Test-Padrao`, which forces
`CultureInvariant`.

## 7. Resolve by marker, not by path

> **Code that locates a folder by a fixed path breaks at the first
> reorganisation. Code that locates it by a marker survives.**

The comparison is by **exact path component**, never by prefix:

```powershell
return ($Path.Split([char]92, [char]47) -contains $script:MarcaEcossistema)
```

The `Split` on `[char]92` instead of a written character class is deliberate. The
previous version used `[\\/]`, a tooling layer ate one backslash leaving `[/]`,
and the function returned false for **every** Windows path.

> **The parse passed, the regex compiled, and the protection simply did not
> happen.** A negative test caught it. A positive test cannot see a protection
> that stopped protecting.

The same family: **an empty alternative in an alternation matches anything.**
`'Spotify' -match 'a||b'` is `True`. In a protection list that becomes "protect
everything", every negative test passes, and nothing breaks — the hardest failure
mode in this code to notice.

## 8. An empty configuration value must not become a wildcard

`-like ('*' + $value + '*')` with an empty value is `-like '**'`, which matches
every path. Measured: every large file came back as KEEP with the reason *"it is
in the clinical folder"* — a wrong verdict delivered confidently, with a reason
that was false, and no error line.

`Test-DentroDaPasta` short-circuits on empty **before** building the wildcard, and
compares by path component rather than text prefix — because the prefix `C:\DATA`
also matches `C:\DATA_OLD`.

## 9. A check that did not run must say so

`Get-Service -ErrorAction Stop` trips on the first service the user cannot query.
The trip becomes terminating, the `catch` swallows it, and the list comes back
**empty** — which does not look like an error. It looks like "this machine hosts
no services", and the server-role warning disappeared without a word.

> **Silence is read as "everything is fine". It is the cheapest way to lie by
> accident.**

## 10. No third-party infrastructure identifier

This repository is public. `tools/Verificar-Vazamento.ps1` scans by **category** —
the *shape* of a network identifier: private IP, UNC to a named host, domain with
a TLD, email address — and never by a list of names. A list of names in a
versioned file puts into the repository exactly what it exists to keep out.

Its header declares what it does **not** catch: a corporate name without a TLD, a
UNC to a short host, binary files, and a person's name. A scanner that promises
more than it delivers is worse than no scanner, because it produces the feeling of
having checked.

---

## Before delivering any change

```
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Tudo.ps1
```

The blocks are **negative tests**: each one builds the case where the protection
must fail if it is broken. If a block starts failing, read its name before
touching the test.
