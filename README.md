# parashaoly

[![Build Windows exe](https://github.com/ranemk/musor/actions/workflows/build-windows.yml/badge.svg)](https://github.com/ranemk/musor/actions/workflows/build-windows.yml)

Build instructions for the Windows executable.

## Local Windows Build

Install build dependencies:

```powershell
python -m pip install -r requirements-build.txt
```

Build the exe:

```powershell
python -m PyInstaller --clean --noconfirm .\L2DamageMeter.spec
```

Output:

```text
dist\parashaoly.exe
```

## Docker Build

This project builds a Windows `.exe`, so Docker must be switched to **Windows containers**.
Linux containers cannot build this exe with PyInstaller.

Build the image:

```powershell
docker build -f Dockerfile.windows -t parashaoly-build .
```

Copy the built exe out of the image:

```powershell
docker create --name parashaoly-out parashaoly-build
New-Item -ItemType Directory -Force .\dist-docker
docker cp parashaoly-out:"C:\app\dist\parashaoly.exe" ".\dist-docker\parashaoly.exe"
docker rm parashaoly-out
```

Output:

```text
dist-docker\parashaoly.exe
```

If the Windows version does not match the base image, try Hyper-V isolation:

```powershell
docker build --isolation=hyperv -f Dockerfile.windows -t parashaoly-build .
```
