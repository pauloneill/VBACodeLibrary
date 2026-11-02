# Incidents and reverts log

Every time something bad slips through (including a bug found in a project that imported this code), add an entry.
This becomes your test set for the review prompts: after you change `REVIEW.md` or a prompt, re-run the reviewer on these old diffs and see whether it would now catch them.

| Date | Commit / PR (bad) | Fix / revert | What went wrong | Which layer should have caught it | Change made so it won't recur |
|------|-------------------|--------------|-----------------|-----------------------------------|-------------------------------|
|      |                   |              |                 |                                   |                               |

## Layers (for the "which layer" column)
1. Pre-PR verification (`sh scripts/check.sh`, importing and compiling in the host application, trying it)
2. AI review
3. Security review
4. CI checks
5. Human (my own) review
