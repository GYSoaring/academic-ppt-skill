# Security

## Supported Versions

Security fixes target the latest release on `main`.

## Reporting

Please report a vulnerability privately through GitHub's security advisory flow instead of opening a public issue with exploit details.

## Scope

This project processes local documents and images. It does not upload materials by itself. Be careful with:

- Untrusted PDF, DOCX, PPTX, or image files; parsing libraries can have vulnerabilities.
- Theme JSON that points to arbitrary local font or image paths.
- Running generated commands without reviewing them.

The skill never needs an API key to render a deck. Do not commit credentials, private manuscripts, or unpublished figures.
