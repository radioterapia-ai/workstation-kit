## What changes, and why it was wrong before

<!-- The reason is the part that stops the defect coming back. -->

## Gate

- [ ] `powershell -NoProfile -ExecutionPolicy Bypass -File tools\Testar-Tudo.ps1` — exit 0
- [ ] Nothing here requires administrator
- [ ] Nothing here survives a restart, except files the user chose to delete
- [ ] A protection-list entry, if any, has a test case that was written **first**
- [ ] No hard-coded drive letter, absolute path, IP address or host name

Paste the last lines of the gate output:

```

```

## Tested on

<!-- Windows version, UI language, interface language selected. -->

## What you did not test

<!-- Say it. "Not observed against a machine with a dead network mapping" is a
     useful sentence; silence is not. -->
