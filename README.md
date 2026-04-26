# parashaoly

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

## GitHub Actions Build

GitHub Actions builds the exe automatically on GitHub.

Workflow file:

```text
.github\workflows\build-windows.yml
```

It runs when code is pushed to `main`, and it can also be started manually from the GitHub **Actions** tab.

To download the exe:

1. Open the repo on GitHub.
2. Click **Actions**.
3. Open **Build Windows exe**.
4. Click the latest successful run.
5. Download the **parashaoly-windows** artifact.

To start a build manually:

1. Open **Actions**.
2. Click **Build Windows exe**.
3. Click **Run workflow**.
4. Choose `main`.
5. Click **Run workflow** again.
