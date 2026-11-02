## What and why
<!-- One or two sentences. Link the issue / README section if there is one. -->

## Testing
<!-- What did you run, and what did you see? Paste output or a screenshot. -->
- [ ] `sh scripts/check.sh` passes
- [ ] Imported into a clean project and **Debug > Compile** is clean (host and Office version: )
- [ ] Exercised the change in the host application: 
- [ ] Area `README.md` updated if public members changed, and header change log updated

## Risk
<!-- Pick one and say why. -->
- [ ] Low (docs, comments, isolated internal fix)
- [ ] Medium (new component or member, behaviour change, error handling or cleanup change)
- [ ] High (public signature or enum change, `Attribute` lines, `Assert` classes, `Globals.bas`, data deletion, CI)

What could go wrong, and what check would catch it?

## How to revert
<!-- Single revert commit? Would projects that already imported the changed file need to be told or re-imported? -->

## Review status
- [ ] AI review run (`.github/prompts/pre-pr-review.prompt.md`)
- [ ] Security review run (medium/high risk)
- [ ] All findings fixed or closed with a reason
- [ ] I read the Testing / Risk / Revert answers above and they are accurate
