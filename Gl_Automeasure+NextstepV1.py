import time
import json
from pathlib import Path

from pywinauto import Desktop
from pywinauto.keyboard import send_keys


MAIN_TITLE_KEYWORD = "GL_SpectroSoft - Lab"
DIALOG_TITLE_KEYWORD = "Luminous efficacy"
MEASURE_BUTTON_TEXT = "Measure"

POPUP_TIMEOUT_S = 10.0
DIALOG_CLOSE_TIMEOUT_S = 8.0

LATEST_JSON_PATH = Path("latest_psu.json")
GL_HANDSHAKE_PATH = Path("gl_handshake.json")


def read_latest_vip():
    if not LATEST_JSON_PATH.exists():
        raise RuntimeError("No existe latest_psu.json (la rampa debe generar snapshots).")
    data = json.loads(LATEST_JSON_PATH.read_text(encoding="utf-8"))
    v = float(data["voltage_v"])
    i = float(data["current_a"])
    p = float(data["power_w"])
    return v, i, p


def find_main_window():
    wins = Desktop(backend="uia").windows()
    for w in wins:
        title = (w.window_text() or "")
        if MAIN_TITLE_KEYWORD.lower() in title.lower():
            return w
    raise RuntimeError(f"No encontré ventana: '{MAIN_TITLE_KEYWORD}'")


def click_measure(main_win):
    main_win.set_focus()
    try:
        btn = main_win.child_window(title=MEASURE_BUTTON_TEXT, control_type="Button")
        btn.wait("exists enabled visible", timeout=2)
        btn.invoke()
        return
    except Exception:
        # fallback
        send_keys("%m")


def wait_dialog_open():
    t0 = time.time()
    while time.time() - t0 < POPUP_TIMEOUT_S:
        time.sleep(0.15)
        dialogs = Desktop(backend="uia").windows(control_type="Window")
        for d in dialogs:
            title = (d.window_text() or "")
            if DIALOG_TITLE_KEYWORD.lower() in title.lower():
                return d
    raise RuntimeError("No apareció el popup de Luminous efficacy.")


def wait_dialog_closed():
    t0 = time.time()
    while time.time() - t0 < DIALOG_CLOSE_TIMEOUT_S:
        time.sleep(0.15)
        dialogs = Desktop(backend="uia").windows(control_type="Window")
        found = False
        for d in dialogs:
            title = (d.window_text() or "")
            if DIALOG_TITLE_KEYWORD.lower() in title.lower():
                found = True
                break
        if not found:
            return True
    return False


def fill_dialog_voltage_current_power_and_ok(dlg, voltage_v, current_a, power_w):
    dlg.set_focus()

    edits = dlg.descendants(control_type="Edit")
    if len(edits) < 3:
        raise RuntimeError(f"Esperaba 3 Edit, encontré {len(edits)}")

    # Por tu UI: edits[0]=Power, edits[1]=Current, edits[2]=Voltage
    edit_power = edits[0]
    edit_current = edits[1]
    edit_voltage = edits[2]

    Vtxt = f"{voltage_v:.3f}"
    Itxt = f"{current_a:.4f}"
    Ptxt = f"{power_w:.4f}"

    def set_edit(edit, value_str):
        edit.set_focus()
        send_keys("^a{BACKSPACE}")
        edit.type_keys(value_str, with_spaces=True)

    def read_edit(edit) -> str:
        try:
            t = edit.window_text().strip()
            if t:
                return t
        except Exception:
            pass
        try:
            return (edit.get_value() or "").strip()
        except Exception:
            return ""

    # Anti auto-cálculo: 2 pasadas
    for _ in range(2):
        # ORDEN requerido por ti:
        set_edit(edit_voltage, Vtxt)
        set_edit(edit_current, Itxt)
        set_edit(edit_power, Ptxt)

        if read_edit(edit_voltage) == Vtxt and read_edit(edit_current) == Itxt and read_edit(edit_power) == Ptxt:
            break

    # OK inmediato
    try:
        ok_btn = dlg.child_window(title="OK", control_type="Button")
        ok_btn.wait("exists enabled visible", timeout=2)
        ok_btn.invoke()
    except Exception:
        send_keys("{ENTER}")


def do_gl_measure_once():
    v, i, p = read_latest_vip()
    main = find_main_window()
    click_measure(main)
    dlg = wait_dialog_open()
    fill_dialog_voltage_current_power_and_ok(dlg, v, i, p)

    # Confirmación: el popup se cerró (indicador de que “se procesó”)
    if not wait_dialog_closed():
        # Si no se cerró, igual no avanzamos; esto es tu “seguro”
        raise RuntimeError("El popup no se cerró tras OK (no confirmo guardado).")

    return True


def write_handshake(state, step):
    GL_HANDSHAKE_PATH.write_text(json.dumps({
        "state": state,
        "step": int(step),
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
    }, indent=2), encoding="utf-8")


def main():
    print("GL AutoMeasure + Auto NEXT (file handshake)")
    print(f"- watching: {GL_HANDSHAKE_PATH.resolve()}")
    print(f"- reading:  {LATEST_JSON_PATH.resolve()}")

    last_done_step = None

    while True:
        try:
            if not GL_HANDSHAKE_PATH.exists():
                time.sleep(0.2)
                continue

            h = json.loads(GL_HANDSHAKE_PATH.read_text(encoding="utf-8"))
            state = h.get("state", "")
            step = int(h.get("step", -1))

            if state == "HOLD" and step > 0:
                # evita repetir si el archivo no cambia
                if last_done_step == step:
                    time.sleep(0.2)
                    continue

                print(f"[HOLD] step {step} -> measuring in GL…")
                do_gl_measure_once()
                write_handshake("DONE", step)
                last_done_step = step
                print(f"[DONE] step {step} -> ramp can proceed.")

            time.sleep(0.2)

        except Exception as e:
            print(f"[ERROR] {e}")
            time.sleep(0.5)


if __name__ == "__main__":
    main()
