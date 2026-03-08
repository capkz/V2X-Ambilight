---
name: v2x-protocol
description: Reference guide for the Creative Sound Blaster Katana V2X serial LED protocol. Use when modifying KatanaDevice.cs, adding new LED effects, debugging serial communication, or answering questions about the V2X binary protocol.
---

# Creative Sound Blaster Katana V2X — LED Protocol Reference

**Source:** Reverse-engineered from `CTCDC.dll` via [v2x-ctl](https://git.dog/xx/v2x) (Rust, MIT) and [blog.nns.ee/2026/02/20/katana-v2x-re](https://blog.nns.ee/2026/02/20/katana-v2x-re).

---

## Device

| Field     | Value                          |
|-----------|--------------------------------|
| VID       | `0x041E`                       |
| PID       | `0x3283`                       |
| Transport | USB CDC serial (COM port)      |
| Baud      | 115200                         |
| Protocol  | Text auth → binary command mode |

---

## Auth Sequence (must complete before any binary commands)

1. Send `whoareyou\r\n` — always safe; works in both text and binary mode
2. Device replies with: `whoareyou [h0 h1] [pid0 pid1] [32-byte nonce]` (45+ bytes)
3. Build AES-256-GCM key: `[h0, h1, KEY_DATA(28 bytes), pid0, pid1]`
4. Encrypt nonce with random 12-byte IV: `AesGcm.Encrypt(iv[..12], nonce, ciphertext, tag)`
5. Send: `unlock` + iv(16) + ciphertext(32) + tag(16) + `\r\n`
6. Device replies `unlock_OK` → auth done
7. Send `SW_MODE1\r\n` → switches to binary command mode
8. Send binary ping `5A 03 00` to confirm

**Shortcuts (device already authed):**
- If `whoareyou` reply starts with `0x5A` → already in binary mode, skip auth
- If reply is `Unknown command` or `unlock_OK` → already authenticated

---

## Binary Packet Format

```
5A  3A  [len]  [subcmd]  [payload...]
```

- `0x5A` = frame marker
- `0x3A` = OP_LIGHT (lighting opcode; only opcode used for LEDs)
- `[len]` = count of bytes that follow (subcmd + payload)
- All multi-byte integers are little-endian

---

## Sub-commands (SUB_*)

| Constant         | Byte   | Purpose                        |
|------------------|--------|--------------------------------|
| SUB_POWER        | `0x25` | Enable / disable lighting      |
| SUB_BRIGHTNESS   | `0x27` | Global brightness (0–255)      |
| SUB_MODE         | `0x29` | Animation / effect mode        |
| SUB_PARAM        | `0x2B` | Colors, speed, direction, etc. |
| SUB_COLOR_COUNT  | `0x37` | Number of active color slots   |

---

## Mode Values (SUB_MODE, byte `0x29`)

```
5A 3A 03 29 00 [mode]
```

| `[mode]` | Name     | Description                                            |
|----------|----------|--------------------------------------------------------|
| `0x00`   | Off      | Lights off                                             |
| `0x01`   | Cycle    | Rainbow hue rotation through all colors                |
| `0x03`   | Solo/Mood| **Static (count=1) or 7-zone ambilight (count=7)**     |
| `0x04`   | Wave     | Colors travel across the bar left→right                |
| `0x05`   | Pulsate  | Breathing / pulsating brightness                       |
| `0x07`   | Morph    | Smooth color morphing                                  |
| `0x08`   | Aurora   | Aurora borealis sweep effect                           |

**For ambilight use:** `0x03` with color_count=7.
Solo (static 1-color) = `0x03` with color_count=1.

---

## SUB_PARAM Parameters (byte `0x2B`)

All SUB_PARAM packets: `5A 3A [len] 2B 00 [param_id] [value...]`

| `[param_id]` | Name      | Value format          | Example packet                          |
|--------------|-----------|-----------------------|-----------------------------------------|
| `0x01`       | COLORS    | `count` + N×4 ABGR   | See color packets below                 |
| `0x03`       | SPEED     | u16LE period in ms    | `5A 3A 05 2B 00 03 01 00` = 1 ms       |
| `0x04`       | DIRECTION | u16LE (01=L→R, 00=R→L)| `5A 3A 05 2B 00 04 01 00`              |

---

## Color Packets

**Color format:** `[0xFF, B, G, R]` — ABGR, 4 bytes, alpha always `0xFF`.

### Set 7 zones (ambilight — current app packet):
```
5A 3A 20 2B 00 01 01 [zone0: FF B G R] [zone1: FF B G R] ... [zone6: FF B G R]
```
- `0x20` = len (32 bytes follow)
- `01` = count=7? No — the byte after `01 01` is the first zone. The `01 01` is PARAM_COLORS + implicit count from SUB_COLOR_COUNT.

### Set 1 static color (Solo mode):
```
5A 3A 08 2B 00 01 01 FF BB GG RR
```

---

## Initialization Sequence (current app)

```
SW_MODE1\r\n                          ← switch to binary mode
5A 03 00                              ← ping
5A 3A 02 25 01                        ← lighting on
5A 3A 03 37 00 07                     ← 7 color slots
5A 3A 03 29 00 03                     ← Mood mode (7-zone)
5A 3A 05 2B 00 03 01 00               ← transition speed = 1 ms (instant)
```

Then per-frame (20 fps):
```
5A 3A 20 2B 00 01 01 [7 × FF B G R]  ← set 7 zone colors
```

---

## Key Gotchas

- **Always send `whoareyou` before any binary bytes.** Sending binary first corrupts the text parser and auth fails.
- **Device retains binary mode** after Creative app has authed it — detect via `0x5A` probe response.
- `0x00` on SUB_MODE = lights off (not "static").
- `0x03` + count=7 = Mood mode; has built-in transition animation → set SPEED to 1ms to suppress sliding.
- `0x03` + count=1 = Solo mode; truly static single color.
- Zones are ordered left→right on the bar (zone 0 = leftmost).
- COM port is auto-detected via WMI: `VID_041E` + `PID_3283` in `Win32_PnPEntity`.

---

## Quick Reference: Common Packets

```
# Lighting on/off
5A 3A 02 25 01   (on)
5A 3A 02 25 00   (off)

# Brightness (0x80 = 50%)
5A 3A 02 27 80

# Mode
5A 3A 03 29 00 03   (mood/ambilight)
5A 3A 03 29 00 01   (rainbow cycle)
5A 3A 03 29 00 05   (pulsate/breathe)

# Color slots
5A 3A 03 37 00 07   (7 zones)
5A 3A 03 37 00 01   (1 zone)

# Transition speed (1ms = instant)
5A 3A 05 2B 00 03 01 00

# Set all 7 zones to red (FF 00 00)
5A 3A 20 2B 00 01 01
  FF 00 00 FF   FF 00 00 FF   FF 00 00 FF   FF 00 00 FF
  FF 00 00 FF   FF 00 00 FF   FF 00 00 FF
```
