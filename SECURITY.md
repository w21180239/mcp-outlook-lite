# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 2.x     | Yes                |
| 1.x     | No                 |

## Reporting a Vulnerability

Please use GitHub's **private vulnerability reporting** feature to report security issues. You can find it under **Settings > Security > Private vulnerability reporting** on the repository page.

Do **not** open a public issue for security vulnerabilities.

### What to Include

- A clear description of the vulnerability.
- Step-by-step instructions to reproduce the issue.
- An assessment of the potential impact (e.g., data exposure, privilege escalation, token leakage).
- Any relevant logs, screenshots, or proof-of-concept code.

## Response Timeline

- **Acknowledgement**: within 48 hours of the report.
- **Status update**: within 7 days, including an initial assessment and expected next steps.
- **Resolution**: timeline communicated on a case-by-case basis depending on severity and complexity.

## Scope

### In Scope

- Authentication and authorization flows.
- Token handling and storage.
- Input validation and sanitization.
- Interactions with the Microsoft Graph API initiated by this project.

### Out of Scope

- Vulnerabilities in Azure Active Directory itself.
- Bugs in the Microsoft Graph API.
- Issues in third-party dependencies (please report these to the respective maintainers, though we appreciate a heads-up).

## Responsible Disclosure

We ask that you give us a reasonable amount of time to address the issue before making any information public. We are committed to working with security researchers and will credit reporters in release notes (unless anonymity is preferred).
