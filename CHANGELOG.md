# Changelog

## 1.4.0 — 2026-10-08

- Check the public GitHub version manifest when invoked, cache results, and offer updates without installing them. Offline/network failures leave the academic workflow usable.
- Record review coverage explicitly; unreviewed dimensions cannot appear completed.
- Recompare cached bibliography records, bind historical verdicts to entry fingerprints, and review each citation use independently.
- Preserve year suffixes and citation ambiguity; compare volume/issue values and flag title/author differences.
- Bind final verdicts to manuscript bytes and current screening content. Add `--prepare-final` and relocated-source support.
- Correct appendix wording so partially reviewed entries are not presented as completely verified.

### Migration

Rescreen legacy data without `paper.sha256`. Generate a new empty template with `--prepare-final`, then verify and migrate applicable verdicts; old verdict files without `binding` cannot finalize. Old installations require one manual upgrade to gain update checks.

### Validation

Tests cover review integrity, citation matching, cache behavior, version mismatches, and update failures. Markdown/DOCX/PDF fixture replay is offline with synthetic lookup responses, not a live scholarly validation or a measured accuracy benchmark.
