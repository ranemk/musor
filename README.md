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

This app needs Tesseract OCR installed on Windows, including Russian language data.

The app checks these locations automatically:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
C:\Program Files (x86)\Tesseract-OCR\tesseract.exe
```

If Tesseract is somewhere else, set:

```powershell
$env:TESSERACT_CMD = 'C:\Path\To\tesseract.exe'
```

## Notes

- This is passive screen capture only.
- It does not read game memory.
- It does not inspect packets.
- It does not automate input.

The latest captured chat crop is saved as `last_chat_capture.png`, which is useful for tuning OCR if recognition is poor.

# musor
