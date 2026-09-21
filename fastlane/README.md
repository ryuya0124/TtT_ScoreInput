fastlane documentation
----

# Installation

Make sure you have the latest version of the Xcode command line tools installed:

```sh
xcode-select --install
```

For _fastlane_ installation instructions, see [Installing _fastlane_](https://docs.fastlane.tools/#installing-fastlane)

# Available Actions

## Mac

### mac setup_certificates

```sh
[bundle exec] fastlane mac setup_certificates
```

Install the Developer ID certificate managed by match

### mac build_signed

```sh
[bundle exec] fastlane mac build_signed
```

Build and sign the Apple Silicon application

### mac notarize_release

```sh
[bundle exec] fastlane mac notarize_release
```

Notarize the application and staple the ticket

----

This README.md is auto-generated and will be re-generated every time [_fastlane_](https://fastlane.tools) is run.

More information about _fastlane_ can be found on [fastlane.tools](https://fastlane.tools).

The documentation of _fastlane_ can be found on [docs.fastlane.tools](https://docs.fastlane.tools).
