# VBA Code Library
This is a repository of all the Microsoft Access and Word/Excel VBA code I have written over the years in business
Feel free to use it as is, clone this repo or contribute back

If you like what you see you can Buy me a Coffeee or subscribe to my Patreon page (both coming soon...)

## Development

Work happens on the `prototype` branch; `main` only receives reviewed, CI-green changes.

- Enable the pre-push hook once per clone: `git config core.hooksPath .githooks` (see [.githooks/README.md](.githooks/README.md))
- Static checks: `sh scripts/check.sh` (the same script CI runs). It cannot compile VBA, so also import the changed files and run **Debug > Compile** in the host application
- Conventions for contributors and AI agents: [.github/copilot-instructions.md](.github/copilot-instructions.md); review criteria: [REVIEW.md](REVIEW.md); incident log: [INCIDENTS.md](INCIDENTS.md)
