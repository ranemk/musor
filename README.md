# Lineage 2 Damage Meter

Passive screen parser for Lineage 2 combat chat. It watches the chat box, extracts lines like:

```text
Вы нанесли 52 урона
rane has given 1447 damage of Drakos Hunter.
rane has received 1029 damage from Drakos Hunter.
```

and keeps outgoing damage, incoming damage, plus outgoing rolling DPS for 5, 10, and 30 seconds.

## Run

Use the bundled Python from this Codex workspace:

```powershell
& 'C:\Users\rane8\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' .\l2_damage_meter.py
```

Then:

The main window has two sections:

- **Damage Counter** parses chat, calculates damage, and opens the damage overlay.
- **Window Overlay** selects windows/zones and can open a windows-only overlay without the damage counter.

For damage:

1. Click **Select chat region**.
2. Drag around the Lineage 2 chat log area.
3. Click **Start**.
4. Click **Open damage overlay**.

For servitor skill timers:

1. Use **Start Barrier 30s** for Servitor Barrier.
2. Use **Start Empowerment 60s** for Servitor Empowerment.
3. In the overlay, click the skill icon to restart that timer.
4. Use **TIM** in the overlay title bar, or **Servitor timers: On/Off** in the main window, to hide/show the timer section.

For mirrored windows only:

1. Click **Pick overlay window**, then click the window where the overlay should stay.
2. Click **Pick zone window**, then click the window used for mirrored Z1/Z2 zones.
3. Click **Select zone 1** and **Select zone 2**.
4. Click **Open windows-only overlay**.

The app saves the selected rectangle in `l2_damage_meter_config.json`, so you only need to select it again if the game window moves or changes size.
The parse interval can be changed in the main window. Use values like `1.0` for once per second or `0.2` for five checks per second, then click **Apply interval**.
The mirrored zone refresh speed can be changed with **Overlay FPS**. Try `15` or `30` for smoother mirrored zones.
Use **Mode: Fast** while fighting if the counter misses damage. Use **Mode: Safe** while testing idle behavior or when old visible chat is being re-counted.
When a zone window is picked before selecting zones, the zones are saved relative to that window. The damage overlay position is saved relative to the damage window.

The overlay can be dragged by holding the left mouse button on it. Click its `x` button, or click **Overlay** again in the main window, to hide it.
The overlay also has **Start** and **Reset** buttons for the damage counter.
Use the **DMG** button in the overlay title bar to show or hide the damage counter section.
Use the **Z1** and **Z2** buttons in the overlay to show or hide the two mirrored zones.
Each mirrored zone also has its own zoom button, such as `1.0x`, and an `x` button on the right to hide only that zone.

When you press **Start** or **Reset**, the currently visible chat is used as a baseline, so old visible damage lines are not added again.

## OCR Requirement

This app uses the native Windows OCR API through `win_ocr.ps1`.
No external OCR install is required.

If OCR does not read Russian text, install the Russian OCR/language feature in Windows language settings.

## Docker Build

This project builds a Windows `.exe`, so Docker must be switched to **Windows containers** first.
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

The Docker-built exe will be at:

```text
dist-docker\parashaoly.exe
```

If Docker says the Windows base image cannot be found, switch Docker Desktop from Linux containers to Windows containers.
If the Windows version does not match the base image, try running the build with Hyper-V isolation:

```powershell
docker build --isolation=hyperv -f Dockerfile.windows -t parashaoly-build .
```

## GitHub Actions Build

GitHub Actions is GitHub's automatic builder.
When code is pushed to the `main` branch, GitHub starts a temporary Windows machine, builds `parashaoly.exe`, and saves it as a downloadable artifact.

This repo includes:

```text
.github\workflows\build-windows.yml
```

After pushing to GitHub:

1. Open the repo on GitHub.
2. Click the **Actions** tab.
3. Open **Build Windows exe**.
4. Click the latest run.
5. Download the **parashaoly-windows** artifact.

You can also start a build manually:

1. Open **Actions**.
2. Click **Build Windows exe**.
3. Click **Run workflow**.
4. Choose `main`.
5. Click **Run workflow** again.

The workflow does the same build command as local builds:

```powershell
python -m PyInstaller --clean --noconfirm .\L2DamageMeter.spec
```

## Notes

- This is passive screen capture only.
- It does not read game memory.
- It does not inspect packets.
- It does not automate input.

The latest captured chat crop is saved as `last_chat_capture.png`, which is useful for tuning OCR if recognition is poor.

# musor
