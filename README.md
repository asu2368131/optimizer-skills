# Optimizer Skills

[日本語版](ja/README.md)

An evolving collection of practical skills that help improve software, automations, workflows, and technical systems. Each skill focuses on meaningful improvements, including performance, reliability, maintainability, security, cost efficiency, and developer experience, then turns them into clear and actionable next steps.

The goal is not optimization for its own sake. These skills help you choose improvements that fit the system's purpose, scale, and risk.

## Skills

| Skill | Focus |
| --- | --- |
| [`gas-production-readiness-auditor`](gas-production-readiness-auditor/SKILL.md) | Production-readiness audits for Google Apps Script projects. |

Each skill contains its own guidance on when to use it, what it evaluates, and how it presents results.

## How to use

Install this repository as a skill source for your coding assistant, then ask it to help optimize the system or problem you are working on. For example:

```text
Review this project and identify the highest-impact improvements.
```

```text
Help me make this workflow more reliable and easier to maintain.
```

```text
Analyze this code for performance, security, and operational risks.
```

Choose the skill that best matches the technology and goal. The skill's `SKILL.md` file explains its specific triggers, scope, and output format.

## Install

Clone the repository and add the skill directory or directories you need to the skills location used by your coding assistant. Keep each skill directory and its `SKILL.md` file intact.

```bash
git clone https://github.com/asu2368131/optimizer-skills.git
```

Consult your assistant's documentation for the exact skills directory and installation method.

## Contributing

Contributions are welcome. Keep new skills focused, practical, and documented in a `SKILL.md` file. Explain the problem each skill solves, when it should be used, the improvements it targets, and the format of its output.

## License

This project is licensed under the [MIT License](LICENSE).
