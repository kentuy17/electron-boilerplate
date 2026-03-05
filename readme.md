# electron-boilerplate

> Boilerplate to kickstart creating an app with [Electron](https://github.com/electron/electron)

See [awesome-electron](https://github.com/sindresorhus/awesome-electron) for more useful Electron resources.

See [Caprine](https://github.com/sindresorhus/caprine) for a production app based on this boilerplate.

## Features

- [`electron-builder`](https://www.electron.build) fully set up to create cross-platform builds
- [Builds the app on CI](https://www.electron.build/multi-platform-build.html)
- [Silent auto-updates](https://www.electron.build/auto-update.html)
- App menu that adheres to the system user interface guidelines
- [Config handling](https://github.com/sindresorhus/electron-store)
- [Context menu](https://github.com/sindresorhus/electron-context-menu)
- [User-friendly handling of unhandled errors](https://github.com/sindresorhus/electron-unhandled)
- Easily publish new versions to GitHub Releases
- And much more!

## Getting started

**Click the "Use this template" button.**

Alternatively, create a new directory and then run:

```sh
curl -fsSL https://github.com/sindresorhus/electron-boilerplate/archive/main.tar.gz | tar -xz --strip-components 1
```

There's also a [Yeoman generator](https://github.com/sindresorhus/generator-electron).

---

**Remove everything from here and above**

---

# App Name

> The best app ever

## Install

*macOS 10.13+, Linux, and Windows 7+ are supported (64-bit only).*

**macOS**

[**Download**](https://github.com/user/repo/releases/latest) the `.dmg` file.

**Linux**

[**Download**](https://github.com/user/repo/releases/latest) the `.AppImage` or `.deb` file.

*The AppImage needs to be [made executable](http://discourse.appimage.org/t/how-to-make-an-appimage-executable/80) after download.*

**Windows**

[**Download**](https://github.com/user/repo/releases/latest) the `.exe` file.

---

## Dev

Built with [Electron](https://electronjs.org).

### Prerequisites

- [Node.js](https://nodejs.org/) **20 LTS** (recommended for this Electron 13 template)
- npm 10+
- Python 3 and build tooling for native modules (`sqlite3`)
  - macOS: Xcode Command Line Tools (`xcode-select --install`)
  - Ubuntu/Debian: `sudo apt-get install -y build-essential python3 make g++`
  - Windows: Visual Studio Build Tools (Desktop development with C++)

### Installation

```sh
npm install
```

If native module compilation fails, clean and reinstall dependencies:

```sh
rm -rf node_modules package-lock.json
npm install
```

### Run

```sh
npm install
npm start
```

### Publish

```sh
npm run release
```

After Travis finishes building your app, open the release draft it created and click "Publish".
