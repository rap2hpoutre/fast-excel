# Contributing to FastExcel

Thanks for your interest in improving FastExcel! This guide covers everything you
need to get a change merged.

## Philosophy

FastExcel is a small, fast wrapper around [OpenSpout](https://github.com/openspout/openspout).
The goal is a minimal, intuitive API — not to expose every OpenSpout feature. When
proposing a change, lead with the **use case** so we can keep the surface area small.

## Reporting issues

- **Bugs and feature requests** go in the [issue tracker](https://github.com/rap2hpoutre/fast-excel/issues),
  using the provided templates.
- **Usage questions** ("how do I…?") are best asked on Stack Overflow after checking
  the [README](README.md). The tracker is reserved for bugs and feature requests.
- Always **search existing issues** first — and include a minimal reproduction.

## Development setup

```bash
git clone https://github.com/rap2hpoutre/fast-excel.git
cd fast-excel
composer install
```

## Running the tests

```bash
composer test
```

All new behavior should come with a test. Tests live in `tests/`; bug fixes that
close an issue are usually added to `tests/IssuesTest.php` as `testIssue<number>()`.

## Code style

Code style is enforced automatically by **StyleCI** on every pull request — if it
flags something, apply the suggested fix. To stay consistent with the existing code:

- No spaces around string concatenation: `$a.$b` (not `$a . $b`).
- Keep multi-line PHPDoc `@param` blocks vertically aligned.
- Prefer native PHP functions over helpers where equivalent (e.g. `str_ends_with()`).

## Submitting a pull request

1. Branch off `master` (fork the repo if you don't have push access).
2. Keep the PR **focused** — one logical change. Smaller PRs get reviewed faster.
3. Add tests and make sure `composer test` passes.
4. Open the PR against `master` and fill in the template.
5. Make sure CI is green (build matrix + StyleCI).

## Supported versions

FastExcel targets **PHP 8.0+** and the actively supported Laravel releases. CI runs
the suite across the PHP versions listed in
[`.github/workflows/tests.yml`](.github/workflows/tests.yml); please make sure your
change works across that range.

Thanks again — every fix and improvement helps! 🙌
