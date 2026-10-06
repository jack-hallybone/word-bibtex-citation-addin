## Guiding question

> What is the simplest, clearest and most maintainable change that satisfies the stated intent, introduces no known security or data-integrity issues, and has sufficient evidence to establish that?

## Rules

- Don't hand over a change with known security or data-integrity issues without my explicit consent.
- Commit your work, but don't add a remote, push or pull. I download the codebase and take it from there.
- This project is all rights reserved. Don't declare an open-source license for it, and don't use dependencies whose licenses conflict with that.
- Comply with all applicable license terms and intellectual-property requirements for every dependency, third-party asset and piece of copied or adapted code used in the project.
- Before adding a dependency, confirm that it's a well-established package on an official registry (not a look-alike name), actively maintained in the last year, and free of known security issues.
- Never commit or hand over secrets or personal data without my explicit consent.
- Maintain the following Docker compose services:
  - `build` builds the project: `docker compose run --rm build`
  - `preview` runs the built project so I can try it: `docker compose up preview`
  - `verify` is the authoritative check covering at least the tests, a license check, a vulnerability audit and secrets scan: `docker compose run --rm verify`
- At handover, give the `verify` result, or say plainly that it couldn't run.
